// ═══════════════════════════════════════════════════════════
//  CONSELHO DE CLASSE — app.js
//  Lê três abas: "Notas", "FALTAS" e "DIGITAL"
//  Cruza pelo trio ESTUDANTE + TURMA + DISCIPLINA
//  DIGITAL: cruza por TURMA + DISCIPLINA (independente de FALTAS)
//  Frequência calculada por bimestre (aulas dadas × faltas)
//  Situação: nota >= 5 = Aprovado, < 5 = Reprovado
//  Frequência < 75% = alerta (não reprova automaticamente)
// ═══════════════════════════════════════════════════════════

const APP = { notas: [], faltas: [], digital: [], charts: {}, turmaSel: null, discSel: null };

const CFG = {
  mediaAprov: 5.0,
  freqMin: 75,
  escola: 'E.E. Professora Célia Vasques Ferrari Duch',
  etapa: 'Ensino Médio',
};

// ── Utilitários ────────────────────────────────────────────

function toNum(v) {
  if (v === null || v === undefined || v === '' || v === '-') return null;
  if (typeof v === 'number' && isFinite(v)) return v;
  const n = parseFloat(String(v).replace(',', '.'));
  return isFinite(n) ? n : null;
}

function unique(arr) { return [...new Set(arr.filter(Boolean))].sort(); }

function iniciais(nome) {
  const p = nome.trim().split(/\s+/);
  return (p[0][0] + (p[1]?.[0] || '')).toUpperCase();
}

function destroyChart(id) {
  if (APP.charts[id]) { APP.charts[id].destroy(); delete APP.charts[id]; }
}

// Distribui valores inteiros de forma que a soma seja SEMPRE 100%
// Usa o método do maior resto (Hamilton/Largest Remainder Method)
function distribuirPct(valores) {
  const total = valores.reduce((a, b) => a + b, 0);
  if (!total) return valores.map(() => 0);
  const exatos = valores.map(v => v / total * 100);
  const floors = exatos.map(v => Math.floor(v));
  let resto = 100 - floors.reduce((a, b) => a + b, 0);
  const indices = exatos
    .map((v, i) => ({ i, rem: v - floors[i] }))
    .sort((a, b) => b.rem - a.rem);
  for (let k = 0; k < resto; k++) floors[indices[k].i]++;
  return floors;
}

// Chave de cruzamento entre abas
function chave(estudante, turma, disc) {
  return `${estudante}|${turma}|${disc}`;
}

// ── Plugins globais de gráficos ────────────────────────────

// Plugin: percentual em cada barra (valor/total da série no bimestre)
const pluginBarPct = {
  id: 'barPct',
  afterDatasetsDraw(chart) {
    const { ctx, data } = chart;
    if (chart.config.type !== 'bar') return;
    const isHoriz = chart.config.options?.indexAxis === 'y';
    // Pre-computar percentagens por coluna usando distribuirPct
    const numCols = data.labels ? data.labels.length : 0;
    const colPcts = [];
    for (let i = 0; i < numCols; i++) {
      const vals = data.datasets.map(ds => ds.data[i] || 0);
      colPcts.push(distribuirPct(vals));
    }
    data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      meta.data.forEach((bar, i) => {
        const val = ds.data[i];
        if (val === null || val === undefined || val === 0) return;
        const pct = colPcts[i][di];
        if (pct < 8) return; // não mostrar em barras muito pequenas
        const { x, y, width, height } = bar;
        ctx.save();
        ctx.fillStyle = '#fff';
        ctx.font = 'bold 11px sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        if (isHoriz) {
          const cx = x - (x - bar.base) / 2;
          if (Math.abs(x - bar.base) > 34) ctx.fillText(pct + '%', cx, y);
        } else {
          if (height > 18) ctx.fillText(pct + '%', x, y + height / 2);
        }
        ctx.restore();
      });
    });
  }
};

// Plugin: valor médio (nota) em barras horizontais de média
const pluginBarValor = {
  id: 'barValor',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'bar') return;
    if (chart.config.options?.indexAxis !== 'y') return;
    if (chart.data.datasets.length !== 1) return; // só para gráficos de média única
    const { ctx, data } = chart;
    data.datasets[0].data.forEach((val, i) => {
      if (!val) return;
      const meta = chart.getDatasetMeta(0);
      const bar = meta.data[i];
      const { x, y } = bar;
      ctx.save();
      ctx.fillStyle = '#374151';
      ctx.font = 'bold 11px sans-serif';
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      ctx.fillText(val.toFixed(1).replace('.', ','), x + 6, y);
      ctx.restore();
    });
  }
};

// Plugin: anotação de variação % nos gráficos de linha
const pluginLinePct = {
  id: 'linePct',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'line') return;
    const { ctx, data, scales } = chart;
    data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      const pts = meta.data.filter((_, i) => ds.data[i] !== null && ds.data[i] !== undefined);
      pts.forEach((pt, pi) => {
        const idx = meta.data.indexOf(pt);
        if (idx === 0) return;
        // Achar o ponto anterior não-nulo
        let prevIdx = idx - 1;
        while (prevIdx >= 0 && (ds.data[prevIdx] === null || ds.data[prevIdx] === undefined)) prevIdx--;
        if (prevIdx < 0) return;
        const prev = ds.data[prevIdx];
        const curr = ds.data[idx];
        if (prev === null || curr === null) return;
        const delta = curr - prev;
        if (Math.abs(delta) < 0.05) return;
        const sign = delta > 0 ? '+' : '';
        const label = sign + delta.toFixed(1).replace('.', ',');
        ctx.save();
        ctx.font = 'bold 10px sans-serif';
        ctx.fillStyle = delta > 0 ? '#057a55' : '#c81e1e';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(label, pt.x, pt.y - 6);
        ctx.restore();
      });
    });
  }
};

// Plugin: percentual nas fatias do doughnut
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
          const val = ds.data[i];
          if (!val) return;
          const pct = pcts[i];
          if (pct < 5) return;
          const mid = arc.startAngle + (arc.endAngle - arc.startAngle) / 2;
          const r = (arc.innerRadius + arc.outerRadius) / 2;
          ctx2.save();
          ctx2.fillStyle = '#fff';
          ctx2.font = 'bold 13px sans-serif';
          ctx2.textAlign = 'center';
          ctx2.textBaseline = 'middle';
          ctx2.fillText(pct + '%', arc.x + r * Math.cos(mid), arc.y + r * Math.sin(mid));
          ctx2.restore();
        });
      });
    }
  };
}

// Opções base reutilizáveis
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
    plugins: { legend: { display: false } },
    scales: { x: { min: 0, max: 10 } }
  }, extra);
}
function optsLinha(extra) {
  return Object.assign({
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: 'top', labels: { font: { size: 12 }, boxWidth: 12 } } },
    scales: { y: { min: 0, max: 10 } }
  }, extra);
}



// Retorna objeto com b1..b4 (null se vazio)
function notasAluno(n) {
  return { b1: toNum(n.b1), b2: toNum(n.b2), b3: toNum(n.b3), b4: toNum(n.b4) };
}

// Média apenas dos bimestres lançados
function calcMedia(n) {
  const vals = [n.b1, n.b2, n.b3, n.b4].filter(v => v !== null);
  if (!vals.length) return null;
  return +(vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2);
}

// Quantos bimestres têm nota lançada
function bimLancados(n) {
  return [n.b1, n.b2, n.b3, n.b4].filter(v => v !== null).length;
}

// Calcula frequência a partir do registro de faltas
// Soma apenas os bimestres com aulas informadas
function calcFreq(f) {
  const pares = [
    { aulas: toNum(f.a1), faltas: toNum(f.f1) },
    { aulas: toNum(f.a2), faltas: toNum(f.f2) },
    { aulas: toNum(f.a3), faltas: toNum(f.f3) },
    { aulas: toNum(f.a4), faltas: toNum(f.f4) },
  ].filter(p => p.aulas !== null && p.aulas > 0);

  if (!pares.length) return null;
  const totalAulas  = pares.reduce((s, p) => s + p.aulas, 0);
  const totalFaltas = pares.reduce((s, p) => s + (p.faltas || 0), 0);
  return {
    aulas: totalAulas,
    faltas: totalFaltas,
    pct: +((1 - totalFaltas / totalAulas) * 100).toFixed(1),
  };
}

// Situação: apenas pela nota. Frequência é alerta separado.
function situacao(media) {
  if (media === null) return '—';
  return media >= CFG.mediaAprov ? 'Aprovado' : 'Reprovado';
}

// Alerta de frequência (<75%)
function alertaFreq(freq) {
  if (!freq) return false;
  return freq.pct < CFG.freqMin;
}

// É parcial (não tem B4)?
function isParcial(n) {
  return toNum(n.b4) === null;
}

// Evolução: diferença entre último e primeiro bimestre lançados
function evolucao(n) {
  const vals = [n.b1, n.b2, n.b3, n.b4].filter(v => v !== null);
  if (vals.length < 2) return null;
  return +(vals[vals.length - 1] - vals[0]).toFixed(1);
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
  const cls = freq.pct >= 85 ? 'freq-ok' : freq.pct >= CFG.freqMin ? 'freq-al' : 'freq-ruim';
  const alerta = freq.pct < CFG.freqMin ? '<span class="alerta-freq"><i class="bi bi-exclamation-triangle-fill"></i> &lt;75%</span>' : '';
  return `<span class="${cls}">${freq.pct.toFixed(1)}%</span>${alerta}`;
}

function fmtSit(media, freq, parcial) {
  const sit = situacao(media);
  const al  = alertaFreq(freq);
  const suf = parcial ? ' <span style="font-size:10px;">(parc.)</span>' : '';
  if (sit === 'Aprovado') {
    const badge = al
      ? `<span class="badge b-al"><i class="bi bi-exclamation-triangle-fill"></i> Aprovado ⚠ freq.</span>`
      : `<span class="badge b-ap"><i class="bi bi-check-circle-fill"></i> Aprovado</span>`;
    return badge + suf;
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

// ── MATERIAL DIGITAL — Funções de cálculo ─────────────────

// Busca registro digital para turma/disciplina
function getDigital(turma, disc) {
  return APP.digital.filter(d => d.TURMA === turma && d.DISCIPLINA === disc);
}

// Calcula percentual de cumprimento digital para uma lista de registros (total)
// Ignora bimestres com previsto = 0 ou null
function calcDigital(registros) {
  let totalPrev = 0, totalConc = 0;
  registros.forEach(r => {
    [1, 2, 3, 4].forEach(b => {
      const prev = toNum(r[`prev${b}`]);
      const conc = toNum(r[`conc${b}`]);
      if (prev && prev > 0) {
        totalPrev += prev;
        totalConc += (conc || 0);
      }
    });
  });
  if (!totalPrev) return null;
  return {
    previsto: totalPrev,
    concluido: totalConc,
    pct: +((totalConc / totalPrev) * 100).toFixed(1)
  };
}

// Calcula percentual de cumprimento digital por bimestre individualmente
// Retorna array de 4 objetos (null se bimestre sem previsto)
function calcDigitalBimestres(registros) {
  return [1, 2, 3, 4].map(b => {
    let prev = 0, conc = 0;
    registros.forEach(r => {
      const p = toNum(r[`prev${b}`]);
      const c = toNum(r[`conc${b}`]);
      if (p && p > 0) { prev += p; conc += (c || 0); }
    });
    if (!prev) return null;
    return {
      bim: b,
      previsto: prev,
      concluido: conc,
      pct: +((conc / prev) * 100).toFixed(1)
    };
  });
}

// HTML de card de progresso digital (roxo) — com breakdown por bimestre
function fmtDigitalCard(turma, disc, registros) {
  const calc = calcDigital(registros);
  if (!calc) return '';
  const pct = Math.min(100, calc.pct);
  const cls = pct >= 80 ? 'digital-prog-ok' : pct >= 50 ? 'digital-prog-al' : 'digital-prog-re';

  const bims = calcDigitalBimestres(registros);
  const nomeBim = ['1º BIM', '2º BIM', '3º BIM', '4º BIM'];
  const bimRows = bims.map((b, i) => {
    if (!b) return '';
    const bPct  = Math.min(100, b.pct);
    const bCls  = bPct >= 80 ? 'digital-prog-ok' : bPct >= 50 ? 'digital-prog-al' : 'digital-prog-re';
    return `
      <div class="digital-bim-row">
        <div class="digital-bim-label">
          <span class="digital-bim-nome">${nomeBim[i]}</span>
          <span class="digital-bim-info">${b.concluido} de ${b.previsto}</span>
        </div>
        <div class="digital-bim-bar-wrap">
          <div class="digital-bim-bar ${bCls}" style="width:${bPct}%"></div>
        </div>
        <span class="digital-bim-pct ${bCls}">${b.pct.toFixed(0)}%</span>
      </div>`;
  }).join('');

  return `
    <div class="digital-card">
      <div class="digital-card-head">
        <span class="digital-card-turma">${turma}</span>
        <span class="digital-card-disc">${disc}</span>
        <span class="digital-pct ${cls}">${calc.pct.toFixed(1)}%</span>
      </div>
      <div class="digital-prog-bar-wrap">
        <div class="digital-prog-bar ${cls}" style="width:${pct}%"></div>
      </div>
      <div class="digital-card-info">${calc.concluido} de ${calc.previsto} aulas concluídas</div>
      <div class="digital-bim-breakdown">${bimRows}</div>
    </div>`;
}

// ── Parser ─────────────────────────────────────────────────

function colIdx(header, ...keys) {
  return header.findIndex(h => keys.some(k => String(h).toUpperCase().includes(k.toUpperCase())));
}

function parseNotasSheet(rows) {
  if (!rows.length) return [];
  const h = rows[0].map(c => String(c || '').trim());
  const iE = colIdx(h, 'ESTUDANTE', 'ALUNO', 'NOME');
  const iT = colIdx(h, 'TURMA');
  const iD = colIdx(h, 'DISCIPLINA');
  const iB1 = colIdx(h, '1º BIMESTRE', '1 BIMESTRE', 'NOTA 1', 'B1', 'BIM1', '1BIM', 'NOTA1');
  const iB2 = colIdx(h, '2º BIMESTRE', '2 BIMESTRE', 'NOTA 2', 'B2', 'BIM2', '2BIM', 'NOTA2');
  const iB3 = colIdx(h, '3º BIMESTRE', '3 BIMESTRE', 'NOTA 3', 'B3', 'BIM3', '3BIM', 'NOTA3');
  const iB4 = colIdx(h, '4º BIMESTRE', '4 BIMESTRE', 'NOTA 4', 'B4', 'BIM4', '4BIM', 'NOTA4');
  if (iE < 0 || iT < 0 || iD < 0) return [];
  return rows.slice(1).map(r => ({
    ESTUDANTE: String(r[iE] || '').trim().toUpperCase(),
    TURMA:     String(r[iT] || '').trim().toUpperCase(),
    DISCIPLINA:String(r[iD] || '').trim().toUpperCase(),
    b1: iB1 >= 0 ? toNum(r[iB1]) : null,
    b2: iB2 >= 0 ? toNum(r[iB2]) : null,
    b3: iB3 >= 0 ? toNum(r[iB3]) : null,
    b4: iB4 >= 0 ? toNum(r[iB4]) : null,
  })).filter(r => r.ESTUDANTE);
}

function parseFaltasSheet(rows) {
  if (!rows.length) return [];
  const h = rows[0].map(c => String(c || '').trim());
  const iE = colIdx(h, 'ESTUDANTE', 'ALUNO', 'NOME');
  const iT = colIdx(h, 'TURMA');
  const iD = colIdx(h, 'DISCIPLINA');
  // Aulas por bimestre — novo formato
  const iA1 = colIdx(h, 'AULAS 1', 'A1', 'AULAS1');
  const iA2 = colIdx(h, 'AULAS 2', 'A2', 'AULAS2');
  const iA3 = colIdx(h, 'AULAS 3', 'A3', 'AULAS3');
  const iA4 = colIdx(h, 'AULAS 4', 'A4', 'AULAS4');
  // Faltas por bimestre — formato existente na planilha
  const iF1 = colIdx(h, 'FALTAS 1', 'F1', 'FALTAS1');
  const iF2 = colIdx(h, 'FALTAS 2', 'F2', 'FALTAS2');
  const iF3 = colIdx(h, 'FALTAS 3', 'F3', 'FALTAS3');
  const iF4 = colIdx(h, 'FALTAS 4', 'F4', 'FALTAS4');
  // Total de aulas (formato antigo — fallback)
  const iTotalAulas = colIdx(h, 'TOTAL DE AULAS', 'TOTAL AULAS');

  if (iE < 0 || iT < 0 || iD < 0) return [];

  const parsed = rows.slice(1).map(r => ({
    ESTUDANTE: String(r[iE] || '').trim().toUpperCase(),
    TURMA:     String(r[iT] || '').trim().toUpperCase(),
    DISCIPLINA:String(r[iD] || '').trim().toUpperCase(),
    a1: iA1 >= 0 ? toNum(r[iA1]) : null,
    a2: iA2 >= 0 ? toNum(r[iA2]) : null,
    a3: iA3 >= 0 ? toNum(r[iA3]) : null,
    a4: iA4 >= 0 ? toNum(r[iA4]) : null,
    f1: iF1 >= 0 ? toNum(r[iF1]) : null,
    f2: iF2 >= 0 ? toNum(r[iF2]) : null,
    f3: iF3 >= 0 ? toNum(r[iF3]) : null,
    f4: iF4 >= 0 ? toNum(r[iF4]) : null,
    totalAulas: iTotalAulas >= 0 ? toNum(r[iTotalAulas]) : null,
  })).filter(r => r.ESTUDANTE);

  // Fallback: se não tem colunas AULAS por bimestre mas tem TOTAL DE AULAS,
  // distribui proporcionalmente pelas faltas que existem
  parsed.forEach(r => {
    if (!r.a1 && !r.a2 && !r.a3 && !r.a4 && r.totalAulas) {
      const bimComFalta = [r.f1, r.f2, r.f3, r.f4].filter(v => v !== null).length || 4;
      const aulasP = +(r.totalAulas / 4).toFixed(0);
      if (r.f1 !== null) r.a1 = aulasP;
      if (r.f2 !== null) r.a2 = aulasP;
      if (r.f3 !== null) r.a3 = aulasP;
      if (r.f4 !== null) r.a4 = aulasP;
    }
  });

  return parsed;
}

// Parser da aba DIGITAL
// Colunas: TURMA, DISCIPLINA, PREVISTO 1º BIM, CONCLUIDO 1º BIM, ..., PREVISTO 4º BIM, CONCLUIDO 4º BIM
function parseDigitalSheet(rows) {
  if (!rows.length) return [];
  const h = rows[0].map(c => String(c || '').trim().toUpperCase());
  const iT = colIdx(h, 'TURMA');
  const iD = colIdx(h, 'DISCIPLINA');
  const iPrev1 = colIdx(h, 'PREVISTO 1', 'PREV1', 'PREV 1');
  const iConc1 = colIdx(h, 'CONCLUIDO 1', 'CONC1', 'CONC 1');
  const iPrev2 = colIdx(h, 'PREVISTO 2', 'PREV2', 'PREV 2');
  const iConc2 = colIdx(h, 'CONCLUIDO 2', 'CONC2', 'CONC 2');
  const iPrev3 = colIdx(h, 'PREVISTO 3', 'PREV3', 'PREV 3');
  const iConc3 = colIdx(h, 'CONCLUIDO 3', 'CONC3', 'CONC 3');
  const iPrev4 = colIdx(h, 'PREVISTO 4', 'PREV4', 'PREV 4');
  const iConc4 = colIdx(h, 'CONCLUIDO 4', 'CONC4', 'CONC 4');

  if (iT < 0 || iD < 0) return [];

  return rows.slice(1).map(r => ({
    TURMA:      String(r[iT] || '').trim().toUpperCase(),
    DISCIPLINA: String(r[iD] || '').trim().toUpperCase(),
    prev1: iPrev1 >= 0 ? toNum(r[iPrev1]) : null,
    conc1: iConc1 >= 0 ? toNum(r[iConc1]) : null,
    prev2: iPrev2 >= 0 ? toNum(r[iPrev2]) : null,
    conc2: iConc2 >= 0 ? toNum(r[iConc2]) : null,
    prev3: iPrev3 >= 0 ? toNum(r[iPrev3]) : null,
    conc3: iConc3 >= 0 ? toNum(r[iConc3]) : null,
    prev4: iPrev4 >= 0 ? toNum(r[iPrev4]) : null,
    conc4: iConc4 >= 0 ? toNum(r[iConc4]) : null,
  })).filter(r => r.TURMA && r.DISCIPLINA);
}

function lerArquivo(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', raw: false });
        let notas = [], faltas = [], digital = [];
        wb.SheetNames.forEach(name => {
          const ws   = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
          const nUp  = name.trim().toUpperCase();
          if (nUp.includes('NOTA') || nUp === 'NOTAS') {
            notas = parseNotasSheet(ws);
          } else if (nUp.includes('FALT')) {
            faltas = parseFaltasSheet(ws);
          } else if (nUp === 'DIGITAL' || nUp.includes('DIGITAL')) {
            digital = parseDigitalSheet(ws);
          } else {
            // Aba desconhecida: tenta como notas se tiver colunas certas
            const tentativa = parseNotasSheet(ws);
            if (tentativa.length) notas.push(...tentativa);
          }
        });
        resolve({ notas, faltas, digital });
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

// ── Drag & Drop ────────────────────────────────────────────
function dragOver(e, id) { e.preventDefault(); document.getElementById(id).classList.add('over'); }
function dragLeave(id)   { document.getElementById(id).classList.remove('over'); }
function drop(e, id) {
  e.preventDefault(); dragLeave(id);
  processarArquivos(e.dataTransfer.files);
}

async function handleUpload(input) {
  if (!input.files.length) return;
  await processarArquivos(input.files);
}

async function processarArquivos(files) {
  loader(true, 'Lendo planilha...');
  const statusEl = document.getElementById('status-up');
  try {
    for (const f of files) {
      const { notas, faltas, digital } = await lerArquivo(f);
      // Mescla sem duplicar
      const chN = new Set(APP.notas.map(r => chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA) + r.b1 + r.b2 + r.b3 + r.b4));
      notas.forEach(r => { if (!chN.has(chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA) + r.b1 + r.b2 + r.b3 + r.b4)) APP.notas.push(r); });
      const chF = new Set(APP.faltas.map(r => chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA)));
      faltas.forEach(r => { if (!chF.has(chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA))) APP.faltas.push(r); });
      const chD = new Set(APP.digital.map(r => `${r.TURMA}|${r.DISCIPLINA}`));
      digital.forEach(r => { if (!chD.has(`${r.TURMA}|${r.DISCIPLINA}`)) APP.digital.push(r); });
    }
    if (!APP.notas.length) throw new Error('Nenhuma nota encontrada. Verifique se a aba se chama "Notas".');
    const turmas = unique(APP.notas.map(r => r.TURMA));
    const discs  = unique(APP.notas.map(r => r.DISCIPLINA));
    const digTxt = APP.digital.length ? ` · ${APP.digital.length} registro(s) digital` : '';
    statusEl.innerHTML = `<span style="color:var(--verde);">✓ ${APP.notas.length} registros de notas · ${APP.faltas.length} de frequência${digTxt} · ${turmas.length} turma(s) · ${discs.length} disciplina(s)</span>`;
    buildNav();
    showSection('geral');
  } catch(err) {
    statusEl.innerHTML = `<span style="color:var(--verm);">Erro: ${err.message}</span>`;
  }
  loader(false);
}

// ── Loader ─────────────────────────────────────────────────
function loader(sim, msg) {
  const el = document.getElementById('loader');
  el.style.display = sim ? 'flex' : 'none';
  if (msg) document.getElementById('loader-msg').textContent = msg;
}

// ── Navegação ──────────────────────────────────────────────
function buildNav() {
  const turmas = unique(APP.notas.map(r => r.TURMA));
  const discs  = unique(APP.notas.map(r => r.DISCIPLINA));
  document.getElementById('nav-turmas').innerHTML = turmas.map(t =>
    `<a class="nav-link" data-t="${t}" onclick="selTurma('${t}')"><i class="bi bi-people"></i> ${t}</a>`).join('');
  document.getElementById('nav-discs').innerHTML = discs.map(d =>
    `<a class="nav-link" data-d="${encodeURIComponent(d)}" onclick="selDisc('${encodeURIComponent(d)}')"><i class="bi bi-journal-bookmark"></i> ${d}</a>`).join('');

  // Atualiza contadores dos accordions
  const countT = document.getElementById('acc-turmas-count');
  const countD = document.getElementById('acc-discs-count');
  if (countT) { countT.textContent = turmas.length; countT.classList.add('visible'); }
  if (countD) { countD.textContent = discs.length; countD.classList.add('visible'); }

  // Abre os accordions automaticamente na primeira carga
  if (turmas.length) openAcc('acc-turmas');
  if (discs.length) openAcc('acc-discs');

  // Mostrar botão PPTX na topbar
  const btnPptxTopbar = document.getElementById('btn-pptx-topbar');
  if (btnPptxTopbar) btnPptxTopbar.style.display = '';

  // Popular filtro de turma na seção medalhistas
  const sel = document.getElementById('filtro-turma-medalhas');
  if (sel) sel.innerHTML = '<option value="">Todas as turmas</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  // Popular filtros da seção digital
  const selDT = document.getElementById('filtro-digital-turma');
  if (selDT) selDT.innerHTML = '<option value="">Todas as turmas</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  const selDD = document.getElementById('filtro-digital-disc');
  if (selDD) selDD.innerHTML = '<option value="">Todas as disciplinas</option>' + discs.map(d => `<option value="${d}">${d}</option>`).join('');
  // Popular filtros da seção Radar Preditivo
  const selRT = document.getElementById('radar-sel-turma');
  if (selRT) selRT.innerHTML = '<option value="">Todas as turmas</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  const selRD = document.getElementById('radar-sel-disc');
  if (selRD) selRD.innerHTML = '<option value="">Todas as disciplinas</option>' + discs.map(d => `<option value="${d}">${d}</option>`).join('');
}

function openAcc(id) {
  const acc = document.getElementById(id);
  if (!acc) return;
  const toggle = acc.querySelector('.sb-acc-toggle');
  const body = acc.querySelector('.sb-acc-body');
  if (!toggle || !body) return;
  toggle.classList.add('open');
  toggle.setAttribute('aria-expanded', 'true');
  body.classList.add('open');
}

function toggleAcc(id) {
  const acc = document.getElementById(id);
  if (!acc) return;
  const toggle = acc.querySelector('.sb-acc-toggle');
  const body = acc.querySelector('.sb-acc-body');
  if (!toggle || !body) return;
  const isOpen = body.classList.contains('open');
  toggle.classList.toggle('open', !isOpen);
  toggle.setAttribute('aria-expanded', String(!isOpen));
  body.classList.toggle('open', !isOpen);
}

function showSection(id) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('#sidebar .nav-link').forEach(l => l.classList.remove('active'));
  document.getElementById('section-' + id)?.classList.add('active');
  document.querySelectorAll(`[data-s="${id}"]`).forEach(l => l.classList.add('active'));
  const titles = {
    upload:    ['Importar dados', 'Selecione os arquivos para começar'],
    geral:     ['Painel geral', 'Visão consolidada'],
    turma:     ['Turma ' + (APP.turmaSel || ''), ''],
    disc:      [APP.discSel || '', 'Análise e intervenções pedagógicas'],
    medalhas:  ['🏆 Medalhistas', 'Ranking por desempenho e frequência'],
    digital:   ['💻 Material Digital', 'Progresso do conteúdo digital por turma e disciplina'],
    radar:     ['🎯 Radar Preditivo', 'Centro de Inteligência Pedagógica — discrepâncias para investigação'],
    relatorios:['Relatórios PDF', 'Gere relatórios para impressão'],
    pptx:     ['Apresentação PPTX', 'Gerar relatório em PowerPoint para reuniões'],
  };
  const [t, s] = titles[id] || ['', ''];
  document.getElementById('tb-title').textContent = t;
  document.getElementById('tb-sub').textContent = s;
  if (id === 'geral') renderGeral();
  if (id === 'turma') renderTurma();
  if (id === 'disc')  renderDisc();
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

// ── Cruzamento notas + faltas ──────────────────────────────
function getFaltas(estudante, turma, disc) {
  return APP.faltas.find(f =>
    f.ESTUDANTE === estudante && f.TURMA === turma && f.DISCIPLINA === disc
  ) || null;
}

// Dados completos de um registro de nota
function dadosReg(n) {
  const f    = getFaltas(n.ESTUDANTE, n.TURMA, n.DISCIPLINA);
  const media = calcMedia(n);
  const freq  = f ? calcFreq(f) : null;
  const sit   = situacao(media);
  const alFreq = alertaFreq(freq);
  const parc  = isParcial(n);
  return { n, f, media, freq, sit, alFreq, parc };
}

// ── PAINEL GERAL ───────────────────────────────────────────
function renderGeral() {
  if (!APP.notas.length) return;
  const turmas = unique(APP.notas.map(r => r.TURMA));
  const discs  = unique(APP.notas.map(r => r.DISCIPLINA));
  const alunos = unique(APP.notas.map(r => r.ESTUDANTE));

  // Situação por aluno (pior disciplina)
  let aprov = 0, reprov = 0, alFreqs = 0;
  alunos.forEach(a => {
    const regs = APP.notas.filter(r => r.ESTUDANTE === a);
    const sits = regs.map(r => dadosReg(r));
    if (sits.some(d => d.sit === 'Reprovado')) reprov++;
    else aprov++;
    if (sits.some(d => d.alFreq)) alFreqs++;
  });

  const medias = APP.notas.map(r => calcMedia(r)).filter(m => m !== null);
  const mg = medias.length ? +(medias.reduce((a,b)=>a+b,0)/medias.length).toFixed(1) : 0;

  const [pctAprovG, pctReprovG] = distribuirPct([aprov, reprov]);
  document.getElementById('mg-geral').innerHTML = `
    <div class="mc bl"><div class="ml">Total de alunos</div><div class="mv c-bl">${alunos.length}</div><div class="ms">${turmas.length} turmas · ${discs.length} disciplinas</div></div>
    <div class="mc gr"><div class="ml">Aprovados</div><div class="mv c-gr">${aprov}</div><div class="ms">${pctAprovG}% do total</div></div>
    <div class="mc vm"><div class="ml">Reprovados</div><div class="mv c-vm">${reprov}</div><div class="ms">${pctReprovG}% do total</div></div>
    <div class="mc am"><div class="ml">Alerta de frequência</div><div class="mv c-am">${alFreqs}</div><div class="ms">freq. &lt; 75%</div></div>
    <div class="mc rx"><div class="ml">Média geral</div><div class="mv c-rx">${mg.toFixed(1)}</div></div>`;

  destroyChart('ch-sit-turmas');
  const apT=[], rpT=[];
  turmas.forEach(t => {
    const al = unique(APP.notas.filter(r=>r.TURMA===t).map(r=>r.ESTUDANTE));
    let ap=0,rp=0;
    al.forEach(a => {
      const regs = APP.notas.filter(r=>r.ESTUDANTE===a&&r.TURMA===t);
      if(regs.some(r=>dadosReg(r).sit==='Reprovado')) rp++; else ap++;
    });
    apT.push(ap); rpT.push(rp);
  });
  APP.charts['ch-sit-turmas'] = new Chart(document.getElementById('ch-sit-turmas'),{
    type:'bar',
    plugins:[pluginBarPct],
    data:{labels:turmas,datasets:[
      {label:'Aprovado',data:apT,backgroundColor:'#057a55'},
      {label:'Reprovado',data:rpT,backgroundColor:'#c81e1e'},
    ]},
    options: optsBarEmpilhado()
  });

  destroyChart('ch-media-disc');
  const mD = discs.map(d=>{const ms=APP.notas.filter(r=>r.DISCIPLINA===d).map(r=>calcMedia(r)).filter(m=>m!==null);return ms.length?+(ms.reduce((a,b)=>a+b,0)/ms.length).toFixed(2):0;});
  APP.charts['ch-media-disc'] = new Chart(document.getElementById('ch-media-disc'),{
    type:'bar',
    plugins:[pluginBarValor],
    data:{labels:discs,datasets:[{label:'Média',data:mD,backgroundColor:mD.map(m=>m>=CFG.mediaAprov?'#057a55':'#c81e1e')}]},
    options: optsBarHoriz()
  });

  // Gráfico Material Digital no painel geral
  if (APP.digital.length) {
    const cardDigGeral = document.getElementById('card-digital-geral');
    if (cardDigGeral) cardDigGeral.style.display = '';
    const discsDig = unique(APP.digital.map(r => r.DISCIPLINA));
    const pctsPorDisc = discsDig.map(d => {
      const regs = APP.digital.filter(r => r.DISCIPLINA === d);
      const c = calcDigital(regs);
      return c ? c.pct : 0;
    });
    destroyChart('ch-digital-geral');
    APP.charts['ch-digital-geral'] = new Chart(document.getElementById('ch-digital-geral'), {
      type: 'bar',
      data: {
        labels: discsDig,
        datasets: [{
          label: '% Cumprimento Digital',
          data: pctsPorDisc,
          backgroundColor: pctsPorDisc.map(p => p >= 80 ? '#6c2bd9' : p >= 50 ? '#a855f7' : '#c084fc'),
          borderRadius: 4,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => `${ctx.raw.toFixed(1)}%` } }
        },
        scales: {
          x: { min: 0, max: 100, ticks: { callback: v => v + '%' } }
        }
      }
    });
  } else {
    const cardDigGeral = document.getElementById('card-digital-geral');
    if (cardDigGeral) cardDigGeral.style.display = 'none';
  }

  document.getElementById('tb-resumo').innerHTML = turmas.map(t => {
    const al = unique(APP.notas.filter(r=>r.TURMA===t).map(r=>r.ESTUDANTE));
    let ap=0,rp=0,alF=0;
    al.forEach(a=>{
      const regs=APP.notas.filter(r=>r.ESTUDANTE===a&&r.TURMA===t);
      const ds=regs.map(r=>dadosReg(r));
      if(ds.some(d=>d.sit==='Reprovado')) rp++; else ap++;
      if(ds.some(d=>d.alFreq)) alF++;
    });
    const ms=APP.notas.filter(r=>r.TURMA===t).map(r=>calcMedia(r)).filter(m=>m!==null);
    const mg2=ms.length?+(ms.reduce((a,b)=>a+b,0)/ms.length).toFixed(1):0;
    return `<tr class="click" onclick="selTurma('${t}')">
      <td><strong>${t}</strong></td><td>${al.length}</td>
      <td><span class="badge b-ap">${ap}</span></td>
      <td><span class="badge b-re">${rp}</span></td>
      <td>${alF>0?`<span class="badge b-al">${alF}</span>`:'<span style="color:var(--cinza)">—</span>'}</td>
      <td class="nota ${mg2>=CFG.mediaAprov?'n-ok':'n-re'}">${mg2.toFixed(1).replace('.',',')}</td>
    </tr>`;
  }).join('');
}

// ── VISÃO TURMA ────────────────────────────────────────────
function renderTurma() {
  const t = APP.turmaSel;
  const todasRegs = APP.notas.filter(r => r.TURMA === t);
  if (!todasRegs.length) return;
  const todasDiscs = unique(todasRegs.map(r=>r.DISCIPLINA));

  document.getElementById('titulo-alunos').textContent = `Alunos — Turma ${t}`;
  document.getElementById('tb-title').textContent = `Turma ${t}`;

  // Popular select de disciplina
  const selD = document.getElementById('filtro-disc-turma');
  if (selD) {
    const prev = selD.value;
    selD.innerHTML = '<option value="">Todas as disciplinas</option>' +
      todasDiscs.map(d=>`<option value="${d}"${d===prev?' selected':''}>${d}</option>`).join('');
  }

  renderTurmaComFiltro();
  renderDigitalTurma(t);
}

// Renderiza cards de material digital na visão de turma
function renderDigitalTurma(turma) {
  const secDigTurma = document.getElementById('digital-turma-section');
  const bodyDigTurma = document.getElementById('digital-turma-body');
  const regsDigTurma = APP.digital.filter(r => r.TURMA === turma);
  if (!regsDigTurma.length) {
    if (secDigTurma) secDigTurma.style.display = 'none';
    return;
  }
  if (secDigTurma) secDigTurma.style.display = '';
  const discs = unique(regsDigTurma.map(r => r.DISCIPLINA));
  bodyDigTurma.innerHTML = `<div class="digital-grid">${discs.map(d => {
    const regs = regsDigTurma.filter(r => r.DISCIPLINA === d);
    return fmtDigitalCard(turma, d, regs);
  }).join('')}</div>`;
}

function renderTurmaComFiltro() {
  const t    = APP.turmaSel;
  const disc = document.getElementById('filtro-disc-turma')?.value || '';
  const regs = APP.notas.filter(r => r.TURMA === t && (!disc || r.DISCIPLINA === disc));
  if (!regs.length) return;

  const discs  = unique(regs.map(r=>r.DISCIPLINA));
  const alunos = unique(regs.map(r=>r.ESTUDANTE));

  let ap=0,rp=0,alF=0;
  alunos.forEach(a=>{
    const ds=regs.filter(r=>r.ESTUDANTE===a).map(r=>dadosReg(r));
    if(ds.some(d=>d.sit==='Reprovado')) rp++; else ap++;
    if(ds.some(d=>d.alFreq)) alF++;
  });
  const ms=regs.map(r=>calcMedia(r)).filter(m=>m!==null);
  const mg=ms.length?+(ms.reduce((a,b)=>a+b,0)/ms.length).toFixed(1):0;
  const discLabel = disc || `${unique(APP.notas.filter(r=>r.TURMA===t).map(r=>r.DISCIPLINA)).length} disciplinas`;

  const [pctApT, pctRpT] = distribuirPct([ap, rp]);
  document.getElementById('mg-turma').innerHTML=`
    <div class="mc bl"><div class="ml">Alunos</div><div class="mv c-bl">${alunos.length}</div><div class="ms">${discLabel}</div></div>
    <div class="mc gr"><div class="ml">Aprovados</div><div class="mv c-gr">${ap}</div><div class="ms">${pctApT}%</div></div>
    <div class="mc vm"><div class="ml">Reprovados</div><div class="mv c-vm">${rp}</div><div class="ms">${pctRpT}%</div></div>
    <div class="mc am"><div class="ml">Alerta freq.</div><div class="mv c-am">${alF}</div></div>
    <div class="mc rx"><div class="ml">Média geral</div><div class="mv c-rx">${mg.toFixed(1).replace('.',',')}</div></div>`;

  destroyChart('ch-turma-disc');
  const mD=discs.map(d=>{const ms2=regs.filter(r=>r.DISCIPLINA===d).map(r=>calcMedia(r)).filter(m=>m!==null);return ms2.length?+(ms2.reduce((a,b)=>a+b,0)/ms2.length).toFixed(2):0;});
  APP.charts['ch-turma-disc']=new Chart(document.getElementById('ch-turma-disc'),{
    type:'bar', plugins:[pluginBarValor],
    data:{labels:discs,datasets:[{label:'Média',data:mD,backgroundColor:mD.map(m=>m>=CFG.mediaAprov?'#057a55':'#c81e1e')}]},
    options: optsBarHoriz()
  });

  destroyChart('ch-turma-sit');
  const total = ap + rp || 1;
  APP.charts['ch-turma-sit'] = new Chart(document.getElementById('ch-turma-sit'), {
    type: 'doughnut', plugins: [makePctLabelPlugin(total)],
    data: { labels: ['Aprovado', 'Reprovado'], datasets: [{ data: [ap, rp], backgroundColor: ['#057a55', '#c81e1e'], borderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '52%',
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 12 }, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => {
          const vals = ctx.chart.data.datasets[ctx.datasetIndex].data;
          const pcts = distribuirPct(vals.map(v => v || 0));
          return `${ctx.label}: ${ctx.raw} (${pcts[ctx.dataIndex]}%)`;
        }}}
      }
    }
  });
  renderAlunos();
}

function renderAlunos() {
  const t    = APP.turmaSel;
  const disc = document.getElementById('filtro-disc-turma')?.value || '';
  const sit  = document.getElementById('filtro-sit').value;
  const bsc  = document.getElementById('busca').value.toLowerCase();
  const regs = APP.notas.filter(r =>
    r.TURMA === t &&
    (!disc || r.DISCIPLINA === disc) &&
    (!bsc  || r.ESTUDANTE.toLowerCase().includes(bsc))
  );
  const linhas = [];
  regs.forEach(r => {
    const d = dadosReg(r);
    if (sit === 'Aprovado'    && d.sit !== 'Aprovado')   return;
    if (sit === 'Reprovado'   && d.sit !== 'Reprovado')  return;
    if (sit === 'Alerta freq.' && !d.alFreq)             return;
    linhas.push(`<tr class="click" onclick="abrirModal('${r.ESTUDANTE.replace(/'/g,"\\'")}','${t}')">
      <td style="font-size:12px;"><strong>${r.ESTUDANTE}</strong></td>
      <td style="font-size:12px;color:var(--cinza);">${r.DISCIPLINA}</td>
      <td>${fmtNota(r.b1)}</td><td>${fmtNota(r.b2)}</td><td>${fmtNota(r.b3)}</td><td>${fmtNota(r.b4)}</td>
      <td>${fmtEvol(r)}</td>
      <td>${fmtMedia(d.media, d.parc)}</td>
      <td>${fmtFreqCell(d.freq)}</td>
      <td>${fmtSit(d.media, d.freq, d.parc)}</td>
    </tr>`);
  });
  document.getElementById('tb-alunos').innerHTML = linhas.join('') ||
    '<tr><td colspan="10" style="text-align:center;color:var(--cinza);padding:20px;">Nenhum registro</td></tr>';
}

// ── VISÃO DISCIPLINA ───────────────────────────────────────
function renderDisc() {
  const disc = APP.discSel;
  const todasRegs = APP.notas.filter(r => r.DISCIPLINA === disc);
  if (!todasRegs.length) return;
  const turmas = unique(todasRegs.map(r=>r.TURMA));
  document.getElementById('titulo-disc').textContent = disc;
  document.getElementById('tb-title').textContent = disc;

  // Popular select de turmas
  const sel = document.getElementById('filtro-turma-disc');
  const prev = sel.value;
  sel.innerHTML = '<option value="">Todas as turmas</option>' +
    turmas.map(t=>`<option value="${t}"${t===prev?' selected':''}>${t}</option>`).join('');

  renderDiscComFiltro();
}

function renderDiscComFiltro() {
  const disc = APP.discSel;
  const ft   = document.getElementById('filtro-turma-disc').value;
  const regs = APP.notas.filter(r => r.DISCIPLINA === disc && (!ft || r.TURMA === ft));
  if (!regs.length) return;

  const turmas = unique(APP.notas.filter(r=>r.DISCIPLINA===disc).map(r=>r.TURMA));
  const label  = document.getElementById('iv-turma-label');
  if (label) label.textContent = ft ? `— ${ft}` : '';

  let ap=0,rp=0,alF=0;
  const ms=[];
  regs.forEach(r=>{
    const d=dadosReg(r);
    if(d.media!==null) ms.push(d.media);
    if(d.sit==='Reprovado') rp++; else if(d.sit==='Aprovado') ap++;
    if(d.alFreq) alF++;
  });
  const mg=ms.length?+(ms.reduce((a,b)=>a+b,0)/ms.length).toFixed(1):0;

  const [pctApD, pctRpD] = distribuirPct([ap, rp]);
  document.getElementById('mg-disc').innerHTML=`
    <div class="mc bl"><div class="ml">Registros</div><div class="mv c-bl">${regs.length}</div><div class="ms">${ft||turmas.length+' turma(s)'}</div></div>
    <div class="mc gr"><div class="ml">Aprovados</div><div class="mv c-gr">${ap}</div><div class="ms">${ms.length?pctApD+'%':''}</div></div>
    <div class="mc vm"><div class="ml">Reprovados</div><div class="mv c-vm">${rp}</div><div class="ms">${ms.length?pctRpD+'%':''}</div></div>
    <div class="mc am"><div class="ml">Alerta freq.</div><div class="mv c-am">${alF}</div></div>
    <div class="mc rx"><div class="ml">Média geral</div><div class="mv c-rx">${mg.toFixed(1).replace('.',',')}</div></div>`;

  // Gráfico evolução — filtra turmas pelo filtro ativo
  destroyChart('ch-disc-evol');
  const cores=['#1a56db','#057a55','#b45309','#c81e1e','#6c2bd9'];
  const turmasGraf = ft ? [ft] : turmas;
  APP.charts['ch-disc-evol']=new Chart(document.getElementById('ch-disc-evol'),{
    type:'line', plugins:[pluginLinePct],
    data:{labels:['1º Bim','2º Bim','3º Bim','4º Bim'],
      datasets:turmasGraf.map((t,i)=>{
        const tr=regs.filter(r=>r.TURMA===t);
        return{label:t,borderColor:cores[i%cores.length],backgroundColor:'transparent',
          pointBackgroundColor:cores[i%cores.length],tension:.3,spanGaps:true,
          data:['b1','b2','b3','b4'].map(b=>{
            const vs=tr.map(r=>r[b]).filter(v=>v!==null);
            return vs.length?+(vs.reduce((a,c)=>a+c,0)/vs.length).toFixed(2):null;
          })};
      })
    },
    options: optsLinha()
  });

  // Gráfico distribuição com percentual
  destroyChart('ch-disc-dist');
  const faixas=['0–2','2–4','4–5','5–7','7–10'];
  const cnt=[0,0,0,0,0];
  ms.forEach(m=>{if(m<2)cnt[0]++;else if(m<4)cnt[1]++;else if(m<5)cnt[2]++;else if(m<7)cnt[3]++;else cnt[4]++;});
  const totalCnt = cnt.reduce((a,b)=>a+b,0)||1;

  // Plugin especial para distribuição: mostra valor E percentual
  const pctsDist = distribuirPct(cnt);
  const pluginDistPct = {
    id:'distPct',
    afterDatasetsDraw(chart){
      const {ctx,data}=chart;
      data.datasets[0].data.forEach((val,i)=>{
        if(!val) return;
        const meta=chart.getDatasetMeta(0);
        const bar=meta.data[i];
        const pct=pctsDist[i];
        ctx.save();
        // Valor no topo da barra
        ctx.fillStyle='#374151';
        ctx.font='bold 11px sans-serif';
        ctx.textAlign='center';
        ctx.textBaseline='bottom';
        ctx.fillText(val, bar.x, bar.y-2);
        // Percentual dentro da barra
        if(bar.height>20){
          ctx.fillStyle='#fff';
          ctx.textBaseline='middle';
          ctx.fillText(pct+'%', bar.x, bar.y+bar.height/2);
        }
        ctx.restore();
      });
    }
  };

  APP.charts['ch-disc-dist']=new Chart(document.getElementById('ch-disc-dist'),{
    type:'bar', plugins:[pluginDistPct],
    data:{labels:faixas,datasets:[{label:'Alunos',data:cnt,
      backgroundColor:['#c81e1e','#c81e1e','#b45309','#057a55','#057a55']}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{y:{ticks:{stepSize:1}}},layout:{padding:{top:18}}}
  });

  renderIntervencoes(regs);
  renderListaDisc();
}

function renderIntervencoes(regs) {
  const criticos  = regs.filter(r=>{const d=dadosReg(r);return d.media!==null&&d.media<3;});
  const reprovN   = regs.filter(r=>{const d=dadosReg(r);return d.media!==null&&d.media>=3&&d.media<CFG.mediaAprov;});
  const alertFreq = regs.filter(r=>dadosReg(r).alFreq);

  let html='';
  if(criticos.length){
    html+=`<div class="iv iv-crit"><i class="bi bi-exclamation-octagon-fill" style="font-size:20px;color:var(--verm);margin-top:2px;flex-shrink:0;"></i>
    <div><div class="iv-title" style="color:var(--verm-esc);">Desempenho crítico — ${criticos.length} aluno(s)</div>
    <div class="iv-desc">${criticos.map(r=>`<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — média ${(calcMedia(r)||0).toFixed(1).replace('.',',')}`).join('<br>')}</div>
    <div class="iv-desc" style="margin-top:5px;">→ Encaminhar para reforço intensivo. Verificar possibilidade de trabalho de compensação de nota e contato com família.</div></div></div>`;
  }
  if(reprovN.length){
    html+=`<div class="iv iv-alert"><i class="bi bi-exclamation-triangle-fill" style="font-size:20px;color:var(--amber);margin-top:2px;flex-shrink:0;"></i>
    <div><div class="iv-title" style="color:var(--amber-esc);">Abaixo da média — ${reprovN.length} aluno(s)</div>
    <div class="iv-desc">${reprovN.map(r=>`<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — média ${(calcMedia(r)||0).toFixed(1).replace('.',',')}`).join('<br>')}</div>
    <div class="iv-desc" style="margin-top:5px;">→ Solicitar trabalho de compensação de nota. Acompanhamento mais próximo e reforço antes do próximo bimestre.</div></div></div>`;
  }
  if(alertFreq.length){
    html+=`<div class="iv iv-freq"><i class="bi bi-calendar-x-fill" style="font-size:20px;color:#f97316;margin-top:2px;flex-shrink:0;"></i>
    <div><div class="iv-title" style="color:#7c2d12;">Frequência abaixo de 75% — ${alertFreq.length} aluno(s)</div>
    <div class="iv-desc">${alertFreq.map(r=>{const freq=calcFreq(getFaltas(r.ESTUDANTE,r.TURMA,r.DISCIPLINA));return`<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — ${freq?freq.pct.toFixed(1)+'%':'-'} de presença`;}).join('<br>')}</div>
    <div class="iv-desc" style="margin-top:5px;">→ Comunicar família. Verificar motivo das faltas e possibilidade de trabalho de compensação de frequência.</div></div></div>`;
  }
  if(!html) html=`<div class="iv iv-ok"><i class="bi bi-check-circle-fill" style="font-size:20px;color:var(--verde);flex-shrink:0;"></i><div><div class="iv-title" style="color:var(--verde-esc);">Sem alertas críticos nesta disciplina</div><div class="iv-desc">Todos os alunos com média ≥ 5,0 e sem alerta de frequência.</div></div></div>`;
  document.getElementById('div-iv').innerHTML=html;
}

function onFiltroTurmaDisc() {
  renderDiscComFiltro();
}

function renderListaDisc() {
  const disc=APP.discSel;
  const ft=document.getElementById('filtro-turma-disc').value;
  const regs=APP.notas.filter(r=>r.DISCIPLINA===disc&&(!ft||r.TURMA===ft));
  document.getElementById('tb-disc').innerHTML=regs.map(r=>{
    const d=dadosReg(r);
    const alerta=d.media!==null&&d.media<3?`<span class="badge b-re">Crítico</span>`:
                 d.media!==null&&d.media<CFG.mediaAprov?`<span class="badge b-al">Comp. nota</span>`:
                 d.alFreq?`<span class="badge b-al">Comp. freq.</span>`:
                 `<span style="color:var(--cinza);font-size:11px;">—</span>`;
    return`<tr class="click" onclick="abrirModal('${r.ESTUDANTE.replace(/'/g,"\\'")}','${r.TURMA}')">
      <td style="font-size:12px;"><strong>${r.ESTUDANTE}</strong></td>
      <td><span class="badge b-info">${r.TURMA}</span></td>
      <td>${fmtNota(r.b1)}</td><td>${fmtNota(r.b2)}</td><td>${fmtNota(r.b3)}</td><td>${fmtNota(r.b4)}</td>
      <td>${fmtMedia(d.media,d.parc)}</td>
      <td>${fmtFreqCell(d.freq)}</td>
      <td>${fmtSit(d.media,d.freq,d.parc)}</td>
      <td>${alerta}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="10" style="text-align:center;color:var(--cinza);padding:20px;">Nenhum registro</td></tr>';
}

// ── SEÇÃO MATERIAL DIGITAL ─────────────────────────────────
function renderDigital() {
  if (!APP.digital.length) {
    document.getElementById('mg-digital').innerHTML = `
      <div class="mc" style="grid-column:1/-1;border-left:4px solid var(--roxo);">
        <div class="ml">Sem dados de material digital</div>
        <div class="mv c-rx" style="font-size:14px;font-weight:500;">Importe uma planilha com a aba <strong>DIGITAL</strong> para visualizar o progresso</div>
      </div>`;
    document.getElementById('digital-cards-body').innerHTML = '';
    const cvs = document.getElementById('ch-digital-disc');
    if (cvs) cvs.parentElement.parentElement.style.display = 'none';
    return;
  }

  const ft = document.getElementById('filtro-digital-turma')?.value || '';
  const fd = document.getElementById('filtro-digital-disc')?.value || '';
  const regs = APP.digital.filter(r => (!ft || r.TURMA === ft) && (!fd || r.DISCIPLINA === fd));

  // Métricas globais
  const calcTotal = calcDigital(regs);
  const turmasDig = unique(regs.map(r => r.TURMA));
  const discsDig  = unique(regs.map(r => r.DISCIPLINA));
  const pctGlobal = calcTotal ? calcTotal.pct : 0;
  const pctCls = pctGlobal >= 80 ? 'c-gr' : pctGlobal >= 50 ? 'c-am' : 'c-vm';

  document.getElementById('mg-digital').innerHTML = `
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Cumprimento geral</div><div class="mv ${pctCls}">${calcTotal ? calcTotal.pct.toFixed(1) + '%' : '—'}</div><div class="ms">${calcTotal ? calcTotal.concluido + ' de ' + calcTotal.previsto + ' aulas' : ''}</div></div>
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Turmas</div><div class="mv c-rx">${turmasDig.length}</div><div class="ms">com material digital</div></div>
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Disciplinas</div><div class="mv c-rx">${discsDig.length}</div><div class="ms">no filtro atual</div></div>
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Registros</div><div class="mv c-rx">${regs.length}</div><div class="ms">turma × disciplina</div></div>`;

  // Gráfico de barras por disciplina
  const cvs = document.getElementById('ch-digital-disc');
  if (cvs) cvs.parentElement.parentElement.style.display = '';
  const pctsPorDisc = discsDig.map(d => {
    const r = regs.filter(x => x.DISCIPLINA === d);
    const c = calcDigital(r);
    return c ? c.pct : 0;
  });
  destroyChart('ch-digital-disc');
  APP.charts['ch-digital-disc'] = new Chart(document.getElementById('ch-digital-disc'), {
    type: 'bar',
    data: {
      labels: discsDig,
      datasets: [{
        label: '% Cumprimento',
        data: pctsPorDisc,
        backgroundColor: pctsPorDisc.map(p => p >= 80 ? '#6c2bd9' : p >= 50 ? '#a855f7' : '#c084fc'),
        borderRadius: 4,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => `${ctx.raw.toFixed(1)}%` } }
      },
      scales: { x: { min: 0, max: 100, ticks: { callback: v => v + '%' } } }
    }
  });

  // Cards de progresso
  const body = document.getElementById('digital-cards-body');
  if (!regs.length) {
    body.innerHTML = '<p style="color:var(--cinza);font-size:13px;">Nenhum dado encontrado para o filtro selecionado.</p>';
    return;
  }

  // Agrupa por turma
  const grupos = {};
  regs.forEach(r => {
    if (!grupos[r.TURMA]) grupos[r.TURMA] = {};
    if (!grupos[r.TURMA][r.DISCIPLINA]) grupos[r.TURMA][r.DISCIPLINA] = [];
    grupos[r.TURMA][r.DISCIPLINA].push(r);
  });

  let html = '';
  Object.keys(grupos).sort().forEach(turma => {
    html += `<div style="margin-bottom:18px;"><div style="font-size:13px;font-weight:700;color:var(--roxo-esc);margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid var(--roxo-cl);">Turma ${turma}</div>`;
    html += `<div class="digital-grid">`;
    Object.keys(grupos[turma]).sort().forEach(disc => {
      html += fmtDigitalCard(turma, disc, grupos[turma][disc]);
    });
    html += `</div></div>`;
  });
  body.innerHTML = html;
}

// ── MODAL ALUNO ────────────────────────────────────────────
function abrirModal(nome, turma) {
  const regs=APP.notas.filter(r=>r.ESTUDANTE===nome&&r.TURMA===turma);
  if(!regs.length) return;
  document.getElementById('m-av').textContent=iniciais(nome);
  document.getElementById('m-nome').textContent=`${nome} — ${turma}`;

  const todos=regs.map(r=>dadosReg(r));
  const temReprov=todos.some(d=>d.sit==='Reprovado');
  const temAlFreq=todos.some(d=>d.alFreq);
  const sitEl=document.getElementById('m-sit');
  sitEl.className='badge '+(temReprov?'b-re':temAlFreq?'b-al':'b-ap');
  sitEl.textContent=temReprov?'Reprovado em disciplina(s)':temAlFreq?'Aprovado ⚠ freq.':'Aprovado';

  const cores=['#1a56db','#057a55','#b45309','#c81e1e','#6c2bd9','#0e9f6e'];

  let tabela=`<table style="width:100%;border-collapse:collapse;font-size:12px;">
    <thead><tr style="background:#f9fafb;">
      <th style="padding:8px 10px;text-align:left;border-bottom:1px solid #e5e7eb;font-size:11px;">Disciplina</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B1</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B2</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B3</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B4</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">Média</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">Frequência</th>
      <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">Situação</th>
    </tr></thead><tbody>`;
  regs.forEach(r=>{
    const d=dadosReg(r);
    tabela+=`<tr>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;font-weight:600;">${r.DISCIPLINA}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(r.b1)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(r.b2)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(r.b3)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(r.b4)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtMedia(d.media,d.parc)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtFreqCell(d.freq)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtSit(d.media,d.freq,d.parc)}</td>
    </tr>`;
  });
  tabela+=`</tbody></table>`;

  // Cards de material digital no modal
  const regsDigModal = APP.digital.filter(r => r.TURMA === turma);
  let htmlDig = '';
  if (regsDigModal.length) {
    const discsDig = unique(regsDigModal.map(r => r.DISCIPLINA));
    htmlDig = `<p style="font-size:12px;font-weight:600;color:var(--cinza-esc);margin:14px 0 8px;"><i class="bi bi-laptop" style="color:var(--roxo);margin-right:4px;"></i> Material Digital — Turma ${turma}</p>
      <div class="digital-grid">${discsDig.map(d => fmtDigitalCard(turma, d, regsDigModal.filter(r => r.DISCIPLINA === d))).join('')}</div>`;
  }

  const nomeEsc=nome.replace(/'/g,"\\'");
  document.getElementById('m-body').innerHTML=
    `<p style="font-size:12px;font-weight:600;color:var(--cinza-esc);margin-bottom:8px;">Evolução bimestral por disciplina</p>`+
    `<div style="position:relative;height:210px;margin-bottom:18px;"><canvas id="ch-modal-evol"></canvas></div>`+
    `<p style="font-size:12px;font-weight:600;color:var(--cinza-esc);margin:0 0 8px;">Notas e frequência</p>`+
    tabela+
    htmlDig+
    `<div style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end;">
      <button class="btn btn-success btn-sm" onclick="pdfAlunoNome('${nomeEsc}','${turma}')"><i class="bi bi-file-earmark-person"></i> Gerar boletim PDF</button>
    </div>`;

  document.getElementById('modal-aluno').classList.add('open');

  // Gráfico de linhas após DOM estar pronto
  setTimeout(()=>{
    const ctx=document.getElementById('ch-modal-evol');
    if(!ctx) return;
    if(APP.charts['ch-modal-evol']){APP.charts['ch-modal-evol'].destroy();delete APP.charts['ch-modal-evol'];}
    APP.charts['ch-modal-evol']=new Chart(ctx,{
      type:'line',
      plugins:[pluginLinePct],
      data:{
        labels:['1º Bim','2º Bim','3º Bim','4º Bim'],
        datasets:regs.map((r,i)=>({
          label:r.DISCIPLINA,
          borderColor:cores[i%cores.length],
          backgroundColor:'transparent',
          pointBackgroundColor:cores[i%cores.length],
          pointRadius:5,
          pointHoverRadius:7,
          tension:0.3,
          spanGaps:true,
          data:[r.b1,r.b2,r.b3,r.b4],
        }))
      },
      options:{
        responsive:true,maintainAspectRatio:false,
        plugins:{
          legend:{position:'top',labels:{font:{size:11},boxWidth:10,padding:12}},
          tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.raw!==null?Number(c.raw).toFixed(1).replace('.',','):'—'}`}},
        },
        scales:{
          y:{min:0,max:10,ticks:{stepSize:1},grid:{color:'rgba(0,0,0,.05)'},
            afterDraw(chart){
              const ctx2=chart.ctx,yScale=chart.scales.y,xScale=chart.scales.x;
              const y5=yScale.getPixelForValue(5);
              ctx2.save();ctx2.beginPath();
              ctx2.setLineDash([5,4]);ctx2.strokeStyle='rgba(220,38,38,0.35)';ctx2.lineWidth=1;
              ctx2.moveTo(xScale.left,y5);ctx2.lineTo(xScale.right,y5);
              ctx2.stroke();ctx2.restore();
            }
          },
          x:{grid:{color:'rgba(0,0,0,.04)'}}
        }
      }
    });
  },60);
}
function fecharModal(e){if(e.target===document.getElementById('modal-aluno'))document.getElementById('modal-aluno').classList.remove('open');}

// ── MEDALHISTAS ────────────────────────────────────────────

// Critérios
const MEDALHAS = [
  { tipo: 'ouro',   emoji: '🥇', titulo: 'Ouro',   mediaMin: 9.0, freqMin: 90, cor: '#f59e0b' },
  { tipo: 'prata',  emoji: '🥈', titulo: 'Prata',  mediaMin: 8.0, freqMin: 90, cor: '#94a3b8' },
  { tipo: 'bronze', emoji: '🥉', titulo: 'Bronze', mediaMin: 7.0, freqMin: 90, cor: '#b45309' },
];

// Calcula média geral do aluno em uma turma (média de todas as disciplinas)
function mediaGeralAluno(nome, turma) {
  const regs = APP.notas.filter(r => r.ESTUDANTE === nome && r.TURMA === turma);
  const ms = regs.map(r => calcMedia(r)).filter(m => m !== null);
  if (!ms.length) return null;
  return +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(2);
}

// Calcula frequência geral do aluno em uma turma (média de todas as disciplinas)
function freqGeralAluno(nome, turma) {
  const regs = APP.notas.filter(r => r.ESTUDANTE === nome && r.TURMA === turma);
  let totalAulas = 0, totalFaltas = 0;
  regs.forEach(r => {
    const f = getFaltas(r.ESTUDANTE, r.TURMA, r.DISCIPLINA);
    if (!f) return;
    const freq = calcFreq(f);
    if (freq) { totalAulas += freq.aulas; totalFaltas += freq.faltas; }
  });
  if (!totalAulas) return null;
  return +((1 - totalFaltas / totalAulas) * 100).toFixed(1);
}

function renderMedalhas() {
  const turmaSel = document.getElementById('filtro-turma-medalhas')?.value || '';
  const turmas = turmaSel ? [turmaSel] : unique(APP.notas.map(r => r.TURMA));

  // Construir lista de alunos com média geral e frequência geral
  const candidatos = [];
  turmas.forEach(t => {
    const alunos = unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE));
    alunos.forEach(a => {
      const media = mediaGeralAluno(a, t);
      const freq  = freqGeralAluno(a, t);
      if (media !== null) candidatos.push({ nome: a, turma: t, media, freq: freq || 0 });
    });
  });

  // Métricas do painel
  const ouro   = candidatos.filter(c => c.media >= 9.0 && c.freq >= 90);
  const prata  = candidatos.filter(c => c.media >= 8.0 && c.media < 9.0 && c.freq >= 90);
  const bronze = candidatos.filter(c => c.media >= 7.0 && c.media < 8.0 && c.freq >= 90);

  document.getElementById('mg-medalhas').innerHTML = `
    <div class="mc" style="border-left:4px solid #f59e0b;">
      <div class="ml">🥇 Ouro</div>
      <div class="mv" style="color:#f59e0b;">${ouro.length}</div>
      <div class="ms">média ≥ 9,0 · freq. ≥ 90%</div>
    </div>
    <div class="mc" style="border-left:4px solid #94a3b8;">
      <div class="ml">🥈 Prata</div>
      <div class="mv" style="color:#64748b;">${prata.length}</div>
      <div class="ms">média ≥ 8,0 · freq. ≥ 90%</div>
    </div>
    <div class="mc" style="border-left:4px solid #b45309;">
      <div class="ml">🥉 Bronze</div>
      <div class="mv" style="color:#b45309;">${bronze.length}</div>
      <div class="ms">média ≥ 7,0 · freq. ≥ 90%</div>
    </div>
    <div class="mc bl">
      <div class="ml">Total premiados</div>
      <div class="mv c-bl">${ouro.length + prata.length + bronze.length}</div>
      <div class="ms">de ${candidatos.length} alunos</div>
    </div>`;

  // Ordenar por média decrescente
  const sort = arr => [...arr].sort((a, b) => b.media - a.media || b.freq - a.freq);

  function listaHTML(arr, tipo) {
    if (!arr.length) return `<div class="medal-empty">Nenhum aluno nesta categoria</div>`;
    return `<ul class="medal-list">${sort(arr).map((c, i) => `
      <li class="medal-item" onclick="abrirModal('${c.nome.replace(/'/g,"\\'")}','${c.turma}')">
        <span class="medal-rank">${i + 1}</span>
        <span class="medal-nome">${c.nome} <span class="medal-turma">${c.turma}</span></span>
        <span class="medal-stats">
          <span class="medal-nota">${c.media.toFixed(1).replace('.', ',')}</span>
          <span class="medal-freq">${c.freq.toFixed(0)}%</span>
        </span>
      </li>`).join('')}</ul>`;
  }

  document.getElementById('medalhas-content').innerHTML = `
    <div class="medal-grid">
      <div class="medal-card medal-ouro">
        <div class="medal-header">
          <span class="medal-emoji">🥇</span>
          <div class="medal-info"><div class="medal-title">Medalha de Ouro</div><div class="medal-sub">Média ≥ 9,0 e frequência ≥ 90%</div></div>
          <span class="medal-count">${ouro.length}</span>
        </div>
        ${listaHTML(ouro, 'ouro')}
      </div>
      <div class="medal-card medal-prata">
        <div class="medal-header">
          <span class="medal-emoji">🥈</span>
          <div class="medal-info"><div class="medal-title">Medalha de Prata</div><div class="medal-sub">Média ≥ 8,0 e frequência ≥ 90%</div></div>
          <span class="medal-count">${prata.length}</span>
        </div>
        ${listaHTML(prata, 'prata')}
      </div>
      <div class="medal-card medal-bronze">
        <div class="medal-header">
          <span class="medal-emoji">🥉</span>
          <div class="medal-info"><div class="medal-title">Medalha de Bronze</div><div class="medal-sub">Média ≥ 7,0 e frequência ≥ 90%</div></div>
          <span class="medal-count">${bronze.length}</span>
        </div>
        ${listaHTML(bronze, 'bronze')}
      </div>
    </div>`;
}


function popSelects(){
  const turmas=unique(APP.notas.map(r=>r.TURMA));
  const discs=unique(APP.notas.map(r=>r.DISCIPLINA));
  ['sel-turma-pdf','sel-turma-aluno'].forEach(id=>{document.getElementById(id).innerHTML=turmas.map(t=>`<option value="${t}">${t}</option>`).join('');});
  document.getElementById('sel-disc-pdf').innerHTML=discs.map(d=>`<option value="${d}">${d}</option>`).join('');
  document.getElementById('sel-turma-disc-pdf').innerHTML='<option value="">Todas</option>'+turmas.map(t=>`<option value="${t}">${t}</option>`).join('');
  popAluno();
}
function popAluno(){
  const t=document.getElementById('sel-turma-aluno').value;
  const al=unique(APP.notas.filter(r=>r.TURMA===t).map(r=>r.ESTUDANTE));
  document.getElementById('sel-aluno-pdf').innerHTML=al.map(a=>`<option value="${a}">${a}</option>`).join('');
}

function estilosPDF(){
  return `body{font-family:Arial,sans-serif;padding:22px;color:#111;}
  h1{font-size:17px;margin:0 0 2px;}h2{font-size:12px;color:#6b7280;margin:0 0 14px;}
  .stats{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap;}
  .stat{padding:9px 16px;border-radius:7px;text-align:center;font-size:12px;min-width:80px;}
  table{width:100%;border-collapse:collapse;}
  th{background:#f9fafb;padding:6px 9px;text-align:left;font-size:11px;color:#374151;border-bottom:1px solid #e5e7eb;white-space:nowrap;}
  td{padding:6px 9px;border-bottom:1px solid #f3f4f6;font-size:12px;}
  .ok{color:#057a55;font-weight:700;}.re{color:#c81e1e;font-weight:700;}.am{color:#b45309;font-weight:700;}.ne{color:#9ca3af;}
  .b-ap{background:#def7ec;color:#014737;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;}
  .b-re{background:#fde8e8;color:#771d1d;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;}
  .b-al{background:#fef3c7;color:#78350f;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;}
  @media print{body{padding:0;}}`;
}

function fmtNotaPDF(v){if(v===null)return'<span class="ne">—</span>';return`<span class="${v>=CFG.mediaAprov?'ok':'re'}">${v.toFixed(1).replace('.',',')}</span>`;}
function fmtSitPDF(media,alFreq,parc){
  const p=parc?' (parc.)':'';
  if(media===null)return'—';
  if(media>=CFG.mediaAprov) return alFreq?`<span class="b-al">Aprov. ⚠freq.${p}</span>`:`<span class="b-ap">Aprovado${p}</span>`;
  return`<span class="b-re">Reprovado${p}</span>`;
}

function pdfTurma(){
  const t=document.getElementById('sel-turma-pdf').value;
  if(!t) return alert('Selecione uma turma.');
  const regs=APP.notas.filter(r=>r.TURMA===t);
  const discs=unique(regs.map(r=>r.DISCIPLINA));
  const alunos=unique(regs.map(r=>r.ESTUDANTE));
  let ap=0,rp=0,alF=0;
  alunos.forEach(a=>{const ds=regs.filter(r=>r.ESTUDANTE===a).map(r=>dadosReg(r));if(ds.some(d=>d.sit==='Reprovado'))rp++;else ap++;if(ds.some(d=>d.alFreq))alF++;});
  const rows=alunos.map(a=>{
    return discs.map(d=>{
      const r=regs.find(x=>x.ESTUDANTE===a&&x.DISCIPLINA===d);
      if(!r) return null;
      const dd=dadosReg(r);
      return`<tr><td style="font-weight:600;">${a}</td><td>${d}</td>
        ${['b1','b2','b3','b4'].map(b=>`<td style="text-align:center;">${fmtNotaPDF(r[b])}</td>`).join('')}
        <td style="text-align:center;">${fmtNotaPDF(dd.media)}</td>
        <td style="text-align:center;">${dd.freq?`<span class="${dd.freq.pct<CFG.freqMin?'am':'ok'}">${dd.freq.pct.toFixed(1)}%</span>`:'<span class="ne">—</span>'}</td>
        <td style="text-align:center;">${fmtSitPDF(dd.media,dd.alFreq,dd.parc)}</td></tr>`;
    }).filter(Boolean).join('');
  }).join('');
  const w=window.open('','_blank');
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Turma ${t}</title><style>${estilosPDF()}</style></head><body>
  <h1>${CFG.escola} — Turma ${t}</h1><h2>${CFG.etapa} · Conselho de Classe · ${new Date().toLocaleDateString('pt-BR')}</h2>
  <div class="stats">
    <div class="stat" style="background:#def7ec;color:#014737;"><strong>${ap}</strong><br>Aprovados</div>
    <div class="stat" style="background:#fde8e8;color:#771d1d;"><strong>${rp}</strong><br>Reprovados</div>
    <div class="stat" style="background:#fef3c7;color:#78350f;"><strong>${alF}</strong><br>Alerta freq.</div>
    <div class="stat" style="background:#f3f4f6;color:#374151;"><strong>${alunos.length}</strong><br>Alunos</div>
  </div>
  <table><thead><tr><th>Aluno</th><th>Disciplina</th><th style="text-align:center;">B1</th><th style="text-align:center;">B2</th><th style="text-align:center;">B3</th><th style="text-align:center;">B4</th><th style="text-align:center;">Média</th><th style="text-align:center;">Frequência</th><th>Situação</th></tr></thead>
  <tbody>${rows}</tbody></table>
  <script>window.onload=()=>window.print();<\/script></body></html>`);
  w.document.close();
}

function pdfAluno(){
  const nome=document.getElementById('sel-aluno-pdf').value;
  const turma=document.getElementById('sel-turma-aluno').value;
  if(!nome) return alert('Selecione um aluno.');
  pdfAlunoNome(nome,turma);
}

function pdfAlunoNome(nome,turma){
  const regs=APP.notas.filter(r=>r.ESTUDANTE===nome&&r.TURMA===turma);
  if(!regs.length) return;
  const todos=regs.map(r=>dadosReg(r));
  const temRep=todos.some(d=>d.sit==='Reprovado');
  const temAlF=todos.some(d=>d.alFreq);
  const sitGeral=temRep?'Reprovado em disciplina(s)':temAlF?'Aprovado ⚠ frequência':'Aprovado';
  const sitCls=temRep?'b-re':temAlF?'b-al':'b-ap';
  const rows=regs.map(r=>{
    const d=dadosReg(r);
    return`<tr><td style="font-weight:600;">${r.DISCIPLINA}</td>
      ${['b1','b2','b3','b4'].map(b=>`<td style="text-align:center;">${fmtNotaPDF(r[b])}</td>`).join('')}
      <td style="text-align:center;">${fmtNotaPDF(d.media)}</td>
      <td style="text-align:center;">${d.freq?`<span class="${d.freq.pct<CFG.freqMin?'am':'ok'}">${d.freq.pct.toFixed(1)}% (${d.freq.faltas} falta${d.freq.faltas!==1?'s':''}/${d.freq.aulas} aulas)</span>`:'<span class="ne">—</span>'}</td>
      <td style="text-align:center;">${fmtSitPDF(d.media,d.alFreq,d.parc)}</td>
      <td style="font-size:11px;color:#6b7280;">${d.media!==null&&d.media<CFG.mediaAprov?'Solicitar comp. nota':d.alFreq?'Solicitar comp. freq.':''}</td>
    </tr>`;
  }).join('');
  const w=window.open('','_blank');
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Boletim ${nome}</title><style>${estilosPDF()}
  .header{background:#1a56db;color:#fff;padding:14px 18px;border-radius:8px;margin-bottom:18px;}
  .header h1{color:#fff;font-size:16px;}.header p{font-size:11px;opacity:.85;margin:0;}</style></head><body>
  <div class="header"><h1>${nome}</h1><p>Turma ${turma} · ${CFG.escola} · ${CFG.etapa} · ${new Date().toLocaleDateString('pt-BR')}</p></div>
  <div class="stats">
    <div class="stat" style="background:${temRep?'#fde8e8':temAlF?'#fef3c7':'#def7ec'};color:${temRep?'#771d1d':temAlF?'#78350f':'#014737'};"><strong>${sitGeral}</strong><br>Situação geral</div>
  </div>
  <table><thead><tr><th>Disciplina</th><th style="text-align:center;">B1</th><th style="text-align:center;">B2</th><th style="text-align:center;">B3</th><th style="text-align:center;">B4</th><th style="text-align:center;">Média</th><th>Frequência</th><th>Situação</th><th>Observação</th></tr></thead>
  <tbody>${rows}</tbody></table>
  <script>window.onload=()=>window.print();<\/script></body></html>`);
  w.document.close();
}

function pdfDisc(){
  const disc=document.getElementById('sel-disc-pdf').value;
  const turma=document.getElementById('sel-turma-disc-pdf').value;
  if(!disc) return alert('Selecione uma disciplina.');
  const regs=APP.notas.filter(r=>r.DISCIPLINA===disc&&(!turma||r.TURMA===turma));
  if(!regs.length) return alert('Nenhum dado encontrado.');
  let ap=0,rp=0,alF=0;
  regs.forEach(r=>{const d=dadosReg(r);if(d.sit==='Reprovado')rp++;else if(d.sit==='Aprovado')ap++;if(d.alFreq)alF++;});
  const rows=regs.map(r=>{
    const d=dadosReg(r);
    const obs=d.media!==null&&d.media<CFG.mediaAprov?'Solicitar comp. de nota':d.alFreq?'Solicitar comp. de frequência':'';
    return`<tr><td style="font-weight:600;">${r.ESTUDANTE}</td>
      <td style="text-align:center;"><span class="b-info" style="background:#e8f0fe;color:#1e3a8a;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;">${r.TURMA}</span></td>
      ${['b1','b2','b3','b4'].map(b=>`<td style="text-align:center;">${fmtNotaPDF(r[b])}</td>`).join('')}
      <td style="text-align:center;">${fmtNotaPDF(d.media)}</td>
      <td style="text-align:center;">${d.freq?`<span class="${d.freq.pct<CFG.freqMin?'am':'ok'}">${d.freq.pct.toFixed(1)}%</span>`:'<span class="ne">—</span>'}</td>
      <td style="text-align:center;">${fmtSitPDF(d.media,d.alFreq,d.parc)}</td>
      <td style="font-size:11px;color:#b45309;font-weight:600;">${obs}</td>
    </tr>`;
  }).join('');
  const w=window.open('','_blank');
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>${disc}</title><style>${estilosPDF()}</style></head><body>
  <h1>${disc}${turma?' — Turma '+turma:' — Todas as turmas'}</h1>
  <h2>${CFG.escola} · ${CFG.etapa} · Conselho de Classe · ${new Date().toLocaleDateString('pt-BR')}</h2>
  <div class="stats">
    <div class="stat" style="background:#def7ec;color:#014737;"><strong>${ap}</strong><br>Aprovados</div>
    <div class="stat" style="background:#fde8e8;color:#771d1d;"><strong>${rp}</strong><br>Reprovados</div>
    <div class="stat" style="background:#fef3c7;color:#78350f;"><strong>${alF}</strong><br>Alerta freq.</div>
    <div class="stat" style="background:#f3f4f6;"><strong>${regs.length}</strong><br>Total</div>
  </div>
  <table><thead><tr><th>Aluno</th><th>Turma</th><th style="text-align:center;">B1</th><th style="text-align:center;">B2</th><th style="text-align:center;">B3</th><th style="text-align:center;">B4</th><th style="text-align:center;">Média</th><th style="text-align:center;">Freq.</th><th>Situação</th><th>Observação</th></tr></thead>
  <tbody>${rows}</tbody></table>
  <script>window.onload=()=>window.print();<\/script></body></html>`);
  w.document.close();
}


// ══════════════════════════════════════════════════════════════
//  MÓDULO PPTX — Geração de apresentação PowerPoint
//  Usa PptxGenJS (carregado via CDN)
//  Cores alinhadas à identidade visual do sistema
// ══════════════════════════════════════════════════════════════

// Paleta PPTX — espelha as variáveis CSS do sistema
const PPTX_CORES = {
  roxo:     '6C2BD9',
  roxoEsc:  '4A1D96',
  roxoCl:   'F3E8FF',
  azul:     '1A56DB',
  azulEsc:  '1E3A8A',
  azulCl:   'EFF6FF',
  verde:    '057A55',
  verdeCl:  'DEF7EC',
  verdeEsc: '014737',
  verm:     'C81E1E',
  vermCl:   'FDE8E8',
  amber:    'B45309',
  amberCl:  'FEF3C7',
  cinza:    '6B7280',
  cinzaCl:  'F3F4F6',
  cinzaEsc: '374151',
  branco:   'FFFFFF',
  escuro:   '0F172A',
  fundo:    'F0F4F8',
};

// ── Helpers de slide ────────────────────────────────────────

function pptxCapa(prs, titulo, subtitulo, detalhe) {
  const slide = prs.addSlide();

  // Fundo escuro com gradiente simulado via retângulos
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: PPTX_CORES.escuro },
  });
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: 0.08,
    fill: { color: PPTX_CORES.roxo },
  });
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 3.42, w: '100%', h: 0.08,
    fill: { color: PPTX_CORES.roxo },
  });

  // Ícone decorativo — círculo roxo
  slide.addShape(prs.ShapeType.ellipse, {
    x: 0.45, y: 0.7, w: 0.9, h: 0.9,
    fill: { color: PPTX_CORES.roxo },
    line: { color: PPTX_CORES.roxo },
  });
  slide.addText('📊', {
    x: 0.45, y: 0.7, w: 0.9, h: 0.9,
    align: 'center', valign: 'middle', fontSize: 22,
  });

  // Título principal
  slide.addText(titulo, {
    x: 0.45, y: 1.75, w: 9.1, h: 0.85,
    fontSize: 32, bold: true, color: PPTX_CORES.branco,
    fontFace: 'Calibri',
  });

  // Subtítulo
  slide.addText(subtitulo, {
    x: 0.45, y: 2.55, w: 9.1, h: 0.5,
    fontSize: 16, color: 'A5B4FC',
    fontFace: 'Calibri',
  });

  // Detalhe (data / info)
  slide.addText(detalhe, {
    x: 0.45, y: 3.12, w: 9.1, h: 0.35,
    fontSize: 11, color: '94A3B8',
    fontFace: 'Calibri',
  });

  return slide;
}

function pptxSlideResumoEscola(prs, dados) {
  const slide = prs.addSlide();
  const { turmas, totalAlunos, aprov, reprov, alFreqs, mg } = dados;

  // Cabeçalho roxo
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: 0.65,
    fill: { color: PPTX_CORES.roxo },
  });
  slide.addText('Painel Consolidado — Escola', {
    x: 0.3, y: 0, w: 9.4, h: 0.65,
    fontSize: 16, bold: true, color: PPTX_CORES.branco,
    fontFace: 'Calibri', valign: 'middle',
  });

  // Cards de métricas — 5 cartões em linha
  const cards = [
    { label: 'Total de Alunos', valor: String(totalAlunos), cor: PPTX_CORES.azulCl, corTexto: PPTX_CORES.azulEsc, borda: PPTX_CORES.azul },
    { label: 'Aprovados',       valor: String(aprov),       cor: PPTX_CORES.verdeCl, corTexto: PPTX_CORES.verdeEsc, borda: PPTX_CORES.verde },
    { label: 'Reprovados',      valor: String(reprov),      cor: PPTX_CORES.vermCl,  corTexto: PPTX_CORES.verm,    borda: PPTX_CORES.verm },
    { label: 'Alerta Freq.',    valor: String(alFreqs),     cor: PPTX_CORES.amberCl, corTexto: PPTX_CORES.amber,   borda: PPTX_CORES.amber },
    { label: 'Média Geral',     valor: String(mg.toFixed(1).replace('.', ',')), cor: PPTX_CORES.roxoCl, corTexto: PPTX_CORES.roxoEsc, borda: PPTX_CORES.roxo },
  ];
  const cW = 1.82, cH = 1.0, cY = 0.82, cGap = 0.065;
  cards.forEach((c, i) => {
    const cX = 0.25 + i * (cW + cGap);
    slide.addShape(prs.ShapeType.rect, {
      x: cX, y: cY, w: cW, h: cH,
      fill: { color: c.cor },
      line: { color: c.borda, pt: 1.5 },
      rectRadius: 0.08,
    });
    slide.addText(c.valor, {
      x: cX, y: cY + 0.12, w: cW, h: 0.52,
      fontSize: 26, bold: true, color: c.corTexto,
      align: 'center', fontFace: 'Calibri',
    });
    slide.addText(c.label, {
      x: cX, y: cY + 0.62, w: cW, h: 0.3,
      fontSize: 10, color: PPTX_CORES.cinzaEsc,
      align: 'center', fontFace: 'Calibri',
    });
  });

  // Tabela resumo por turma
  slide.addText('Resumo por Turma', {
    x: 0.3, y: 2.05, w: 9.4, h: 0.32,
    fontSize: 12, bold: true, color: PPTX_CORES.cinzaEsc,
    fontFace: 'Calibri',
  });

  const tabRows = [
    [
      { text: 'Turma',     options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      { text: 'Alunos',    options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      { text: 'Aprovados', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      { text: 'Reprovados',options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      { text: 'Alerta Freq.',options:{ bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      { text: 'Média',     options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
    ],
    ...turmas.map((info, idx) => {
      const bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl;
      const medCor = info.mg >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm;
      return [
        { text: info.turma,   options: { bold: true, fill: bg, align: 'center' } },
        { text: String(info.alunos),   options: { fill: bg, align: 'center' } },
        { text: String(info.aprov),    options: { fill: bg, align: 'center', color: PPTX_CORES.verde, bold: true } },
        { text: String(info.reprov),   options: { fill: bg, align: 'center', color: info.reprov > 0 ? PPTX_CORES.verm : PPTX_CORES.cinzaEsc, bold: info.reprov > 0 } },
        { text: String(info.alFreq),   options: { fill: bg, align: 'center', color: info.alFreq > 0 ? PPTX_CORES.amber : PPTX_CORES.cinzaEsc } },
        { text: info.mg.toFixed(1).replace('.', ','), options: { fill: bg, align: 'center', color: medCor, bold: true } },
      ];
    }),
  ];

  slide.addTable(tabRows, {
    x: 0.3, y: 2.38, w: 9.4, h: Math.min(1.4, 0.28 * (turmas.length + 1)),
    fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc,
    border: { pt: 0.5, color: 'E5E7EB' },
    rowH: 0.28,
    colW: [1.2, 1.0, 1.4, 1.4, 1.4, 1.0],
  });

  return slide;
}

function pptxSlideResumoPorTurma(prs, turma) {
  const slide = prs.addSlide();
  const regs = APP.notas.filter(r => r.TURMA === turma);
  const alunos = unique(regs.map(r => r.ESTUDANTE));
  let ap = 0, rp = 0, alF = 0;
  alunos.forEach(a => {
    const ds = regs.filter(r => r.ESTUDANTE === a).map(r => dadosReg(r));
    if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++;
    if (ds.some(d => d.alFreq)) alF++;
  });
  const ms = regs.map(r => calcMedia(r)).filter(m => m !== null);
  const mg = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0;
  const discs = unique(regs.map(r => r.DISCIPLINA));

  // Cabeçalho azul
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: 0.65,
    fill: { color: PPTX_CORES.azul },
  });
  slide.addText(`Turma ${turma} — Resumo`, {
    x: 0.3, y: 0, w: 9.4, h: 0.65,
    fontSize: 16, bold: true, color: PPTX_CORES.branco,
    fontFace: 'Calibri', valign: 'middle',
  });

  // Mini-cards
  const cards = [
    { label: 'Alunos',    valor: String(alunos.length), cor: PPTX_CORES.azulCl,  borda: PPTX_CORES.azul,  corTexto: PPTX_CORES.azulEsc },
    { label: 'Aprovados', valor: String(ap),            cor: PPTX_CORES.verdeCl, borda: PPTX_CORES.verde, corTexto: PPTX_CORES.verdeEsc },
    { label: 'Reprovados',valor: String(rp),            cor: PPTX_CORES.vermCl,  borda: PPTX_CORES.verm,  corTexto: PPTX_CORES.verm },
    { label: 'Alerta Freq.', valor: String(alF),        cor: PPTX_CORES.amberCl, borda: PPTX_CORES.amber, corTexto: PPTX_CORES.amber },
    { label: 'Média',     valor: mg.toFixed(1).replace('.', ','), cor: PPTX_CORES.roxoCl, borda: PPTX_CORES.roxo, corTexto: PPTX_CORES.roxoEsc },
  ];
  const cW = 1.82, cH = 0.9, cY = 0.82, cGap = 0.065;
  cards.forEach((c, i) => {
    const cX = 0.25 + i * (cW + cGap);
    slide.addShape(prs.ShapeType.rect, { x: cX, y: cY, w: cW, h: cH, fill: { color: c.cor }, line: { color: c.borda, pt: 1.5 }, rectRadius: 0.07 });
    slide.addText(c.valor, { x: cX, y: cY + 0.08, w: cW, h: 0.48, fontSize: 24, bold: true, color: c.corTexto, align: 'center', fontFace: 'Calibri' });
    slide.addText(c.label, { x: cX, y: cY + 0.55, w: cW, h: 0.28, fontSize: 10, color: PPTX_CORES.cinzaEsc, align: 'center', fontFace: 'Calibri' });
  });

  // Médias por disciplina
  slide.addText('Média por disciplina', {
    x: 0.3, y: 1.9, w: 9.4, h: 0.3,
    fontSize: 11, bold: true, color: PPTX_CORES.cinzaEsc, fontFace: 'Calibri',
  });

  const discRows = [
    [
      { text: 'Disciplina', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul } },
      { text: 'Média',      options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
      { text: 'Situação',   options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
    ],
    ...discs.map((d, idx) => {
      const mds = regs.filter(r => r.DISCIPLINA === d).map(r => calcMedia(r)).filter(m => m !== null);
      const mdg = mds.length ? +(mds.reduce((a, b) => a + b, 0) / mds.length).toFixed(2) : null;
      const bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl;
      const cor = mdg === null ? PPTX_CORES.cinza : mdg >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm;
      const sit = mdg === null ? '—' : mdg >= CFG.mediaAprov ? 'Acima da média' : 'Abaixo da média';
      return [
        { text: d, options: { fill: bg, bold: true } },
        { text: mdg !== null ? mdg.toFixed(1).replace('.', ',') : '—', options: { fill: bg, align: 'center', color: cor, bold: true } },
        { text: sit, options: { fill: bg, align: 'center', color: cor } },
      ];
    }),
  ];

  const tabH = Math.min(1.8, 0.28 * (discs.length + 1));
  slide.addTable(discRows, {
    x: 0.3, y: 2.22, w: 5.5, h: tabH,
    fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc,
    border: { pt: 0.5, color: 'E5E7EB' },
    rowH: 0.28,
    colW: [3.2, 1.2, 1.1],
  });

  // Digital na turma
  const regsDigTurma = APP.digital.filter(r => r.TURMA === turma);
  if (regsDigTurma.length) {
    const calcDig = calcDigital(regsDigTurma);
    if (calcDig) {
      slide.addShape(prs.ShapeType.rect, {
        x: 5.95, y: 2.22, w: 3.75, h: 1.5,
        fill: { color: 'FDF4FF' }, line: { color: 'E9D5FF', pt: 1 }, rectRadius: 0.08,
      });
      slide.addText('💻 Material Digital', {
        x: 6.1, y: 2.3, w: 3.5, h: 0.28,
        fontSize: 11, bold: true, color: PPTX_CORES.roxoEsc, fontFace: 'Calibri',
      });
      const digCor = calcDig.pct >= 80 ? PPTX_CORES.verde : calcDig.pct >= 50 ? PPTX_CORES.amber : PPTX_CORES.verm;
      slide.addText(`${calcDig.pct.toFixed(0)}%`, {
        x: 6.1, y: 2.55, w: 3.5, h: 0.52,
        fontSize: 30, bold: true, color: digCor, fontFace: 'Calibri', align: 'center',
      });
      slide.addText(`${calcDig.concluido} de ${calcDig.previsto} aulas concluídas`, {
        x: 6.1, y: 3.05, w: 3.5, h: 0.25,
        fontSize: 10, color: PPTX_CORES.cinza, fontFace: 'Calibri', align: 'center',
      });
    }
  }

  return slide;
}

function pptxSlidesDetalhado(prs, turma, disc) {
  const regs = APP.notas.filter(r =>
    (!turma || r.TURMA === turma) &&
    (!disc  || r.DISCIPLINA === disc)
  );
  if (!regs.length) return;

  const alunos = unique(regs.map(r => r.ESTUDANTE));
  let ap = 0, rp = 0, alF = 0;
  alunos.forEach(a => {
    const ds = regs.filter(r => r.ESTUDANTE === a).map(r => dadosReg(r));
    if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++;
    if (ds.some(d => d.alFreq)) alF++;
  });
  const ms = regs.map(r => calcMedia(r)).filter(m => m !== null);
  const mg = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0;

  // ── Slide de resumo ──────────────────────────────────────
  const slideSumario = prs.addSlide();
  const headerTitulo = [turma, disc].filter(Boolean).join(' · ') || 'Relatório Detalhado';

  slideSumario.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: 0.65,
    fill: { color: PPTX_CORES.roxo },
  });
  slideSumario.addText(`Resumo — ${headerTitulo}`, {
    x: 0.3, y: 0, w: 9.4, h: 0.65,
    fontSize: 16, bold: true, color: PPTX_CORES.branco, fontFace: 'Calibri', valign: 'middle',
  });

  const metCards = [
    { label: 'Total de Alunos', valor: String(alunos.length), cor: PPTX_CORES.azulCl,  borda: PPTX_CORES.azul,  corV: PPTX_CORES.azulEsc },
    { label: 'Aprovados',       valor: String(ap),            cor: PPTX_CORES.verdeCl, borda: PPTX_CORES.verde, corV: PPTX_CORES.verdeEsc },
    { label: 'Reprovados',      valor: String(rp),            cor: PPTX_CORES.vermCl,  borda: PPTX_CORES.verm,  corV: PPTX_CORES.verm },
    { label: 'Alerta Freq.',    valor: String(alF),           cor: PPTX_CORES.amberCl, borda: PPTX_CORES.amber, corV: PPTX_CORES.amber },
    { label: 'Média Geral',     valor: mg.toFixed(1).replace('.', ','), cor: PPTX_CORES.roxoCl, borda: PPTX_CORES.roxo, corV: PPTX_CORES.roxoEsc },
  ];
  const cW = 1.82, cH = 0.95, cY = 0.8, cGap = 0.065;
  metCards.forEach((c, i) => {
    const cX = 0.25 + i * (cW + cGap);
    slideSumario.addShape(prs.ShapeType.rect, { x: cX, y: cY, w: cW, h: cH, fill: { color: c.cor }, line: { color: c.borda, pt: 1.5 }, rectRadius: 0.07 });
    slideSumario.addText(c.valor, { x: cX, y: cY + 0.1, w: cW, h: 0.52, fontSize: 26, bold: true, color: c.corV, align: 'center', fontFace: 'Calibri' });
    slideSumario.addText(c.label, { x: cX, y: cY + 0.62, w: cW, h: 0.26, fontSize: 10, color: PPTX_CORES.cinzaEsc, align: 'center', fontFace: 'Calibri' });
  });

  // Taxa de aprovação
  const [txAp, txRp] = distribuirPct([ap, rp]);
  const barY = 1.95, barX = 0.3, barW = 9.4, barH = 0.32;
  slideSumario.addShape(prs.ShapeType.rect, { x: barX, y: barY, w: barW, h: barH, fill: { color: PPTX_CORES.vermCl }, line: { color: PPTX_CORES.verm, pt: 0.5 } });
  if (txAp > 0) {
    slideSumario.addShape(prs.ShapeType.rect, {
      x: barX, y: barY, w: +(barW * txAp / 100).toFixed(2), h: barH,
      fill: { color: PPTX_CORES.verde }, line: { color: PPTX_CORES.verde, pt: 0 },
    });
  }
  slideSumario.addText(`Aprovados: ${txAp}%`, { x: barX + 0.1, y: barY, w: 3, h: barH, fontSize: 10, bold: true, color: PPTX_CORES.branco, valign: 'middle', fontFace: 'Calibri' });
  slideSumario.addText(`Reprovados: ${txRp}%`, { x: barX + barW - 2.5, y: barY, w: 2.4, h: barH, fontSize: 10, bold: true, color: PPTX_CORES.verm, align: 'right', valign: 'middle', fontFace: 'Calibri' });

  // Tabela de médias por disciplina (se não filtrou disc)
  if (!disc) {
    const discsRegs = unique(regs.map(r => r.DISCIPLINA));
    slideSumario.addText('Médias por disciplina', {
      x: 0.3, y: 2.42, w: 9.4, h: 0.28,
      fontSize: 11, bold: true, color: PPTX_CORES.cinzaEsc, fontFace: 'Calibri',
    });
    const discTabRows = [
      [
        { text: 'Disciplina', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo } },
        { text: 'Média',      options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      ],
      ...discsRegs.map((d, idx) => {
        const mds = regs.filter(r => r.DISCIPLINA === d).map(r => calcMedia(r)).filter(m => m !== null);
        const mdg = mds.length ? +(mds.reduce((a, b) => a + b, 0) / mds.length).toFixed(2) : null;
        const bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl;
        return [
          { text: d, options: { fill: bg, bold: true } },
          { text: mdg !== null ? mdg.toFixed(1).replace('.', ',') : '—', options: { fill: bg, align: 'center', color: mdg !== null && mdg >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm, bold: true } },
        ];
      }),
    ];
    slideSumario.addTable(discTabRows, {
      x: 0.3, y: 2.72, w: 5.2, h: Math.min(1.2, 0.27 * (discsRegs.length + 1)),
      fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc,
      border: { pt: 0.5, color: 'E5E7EB' }, rowH: 0.27, colW: [3.8, 1.4],
    });
  }

  // ── Slides de dados por página de alunos ────────────────
  const ALUNOS_POR_PAGINA = 14;
  const paginas = Math.ceil(regs.length / ALUNOS_POR_PAGINA);

  for (let p = 0; p < paginas; p++) {
    const slideD = prs.addSlide();
    const pageRegs = regs.slice(p * ALUNOS_POR_PAGINA, (p + 1) * ALUNOS_POR_PAGINA);
    const pageLabel = paginas > 1 ? ` (${p + 1}/${paginas})` : '';

    slideD.addShape(prs.ShapeType.rect, {
      x: 0, y: 0, w: '100%', h: 0.55,
      fill: { color: PPTX_CORES.azulEsc },
    });
    slideD.addText(`Lista de Alunos${pageLabel} — ${headerTitulo}`, {
      x: 0.3, y: 0, w: 9.4, h: 0.55,
      fontSize: 13, bold: true, color: PPTX_CORES.branco, fontFace: 'Calibri', valign: 'middle',
    });

    const hasDisc = !!disc;
    const cabecalho = hasDisc
      ? [
          { text: 'Aluno',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul } },
          { text: 'Turma',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'B1',          options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'B2',          options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'B3',          options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'B4',          options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'Média',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'Freq.',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'Situação',    options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
        ]
      : [
          { text: 'Aluno',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul } },
          { text: 'Média',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'Freq.',       options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
          { text: 'Situação',    options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } },
        ];

    const dataRows = pageRegs.map((r, idx) => {
      const d = dadosReg(r);
      const bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl;
      const medCor = d.media === null ? PPTX_CORES.cinza : d.media >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm;
      const freqCor = !d.freq ? PPTX_CORES.cinza : d.freq.pct >= CFG.freqMin ? PPTX_CORES.verde : PPTX_CORES.verm;
      const freqStr = d.freq ? `${d.freq.pct.toFixed(1)}%` : '—';
      const sitStr = d.sit === 'Aprovado' ? (d.alFreq ? 'Apr. ⚠Freq' : 'Aprovado') : d.sit === 'Reprovado' ? 'Reprovado' : '—';
      const sitCor = d.sit === 'Aprovado' ? PPTX_CORES.verde : d.sit === 'Reprovado' ? PPTX_CORES.verm : PPTX_CORES.cinza;

      const notaCell = (v) => ({
        text: v !== null ? v.toFixed(1).replace('.', ',') : '—',
        options: {
          fill: v !== null && v < CFG.mediaAprov ? PPTX_CORES.vermCl : bg,
          align: 'center',
          color: v !== null && v < CFG.mediaAprov ? PPTX_CORES.verm : PPTX_CORES.cinzaEsc,
          bold: v !== null && v < CFG.mediaAprov,
        },
      });

      if (hasDisc) {
        return [
          { text: r.ESTUDANTE, options: { fill: bg, fontSize: 9 } },
          { text: r.TURMA,     options: { fill: bg, align: 'center' } },
          notaCell(r.b1), notaCell(r.b2), notaCell(r.b3), notaCell(r.b4),
          { text: d.media !== null ? d.media.toFixed(1).replace('.', ',') : '—', options: { fill: d.media !== null && d.media < CFG.mediaAprov ? PPTX_CORES.vermCl : bg, align: 'center', color: medCor, bold: true } },
          { text: freqStr, options: { fill: bg, align: 'center', color: freqCor, bold: d.alFreq } },
          { text: sitStr,  options: { fill: bg, align: 'center', color: sitCor, bold: true } },
        ];
      } else {
        return [
          { text: r.ESTUDANTE, options: { fill: bg, fontSize: 9 } },
          { text: d.media !== null ? d.media.toFixed(1).replace('.', ',') : '—', options: { fill: d.media !== null && d.media < CFG.mediaAprov ? PPTX_CORES.vermCl : bg, align: 'center', color: medCor, bold: true } },
          { text: freqStr, options: { fill: bg, align: 'center', color: freqCor, bold: d.alFreq } },
          { text: sitStr,  options: { fill: bg, align: 'center', color: sitCor, bold: true } },
        ];
      }
    });

    const colWidths = hasDisc ? [3.0, 0.7, 0.7, 0.7, 0.7, 0.7, 0.85, 0.8, 1.05] : [6.2, 1.2, 1.2, 1.8];
    const tableW = colWidths.reduce((a, b) => a + b, 0);
    slideD.addTable([cabecalho, ...dataRows], {
      x: 0.3, y: 0.65, w: tableW, h: 3.7,
      fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc,
      border: { pt: 0.5, color: 'E5E7EB' },
      rowH: 0.245,
      colW: colWidths,
    });
  }

  // ── Slide Material Digital (se houver) ──────────────────
  const regsDigFilt = APP.digital.filter(r =>
    (!turma || r.TURMA === turma) &&
    (!disc  || r.DISCIPLINA === disc)
  );
  if (regsDigFilt.length) {
    const slideD2 = prs.addSlide();
    slideD2.addShape(prs.ShapeType.rect, {
      x: 0, y: 0, w: '100%', h: 0.65,
      fill: { color: PPTX_CORES.roxo },
    });
    slideD2.addText('💻 Material Digital — Progresso por Bimestre', {
      x: 0.3, y: 0, w: 9.4, h: 0.65,
      fontSize: 16, bold: true, color: PPTX_CORES.branco, fontFace: 'Calibri', valign: 'middle',
    });

    const discsD = unique(regsDigFilt.map(r => r.DISCIPLINA));
    const turmasD = unique(regsDigFilt.map(r => r.TURMA));
    const combos = [];
    turmasD.forEach(t => discsD.forEach(d => {
      const rr = regsDigFilt.filter(x => x.TURMA === t && x.DISCIPLINA === d);
      if (rr.length) combos.push({ turma: t, disc: d, regs: rr });
    }));

    const digTabRows = [
      [
        { text: 'Turma',      options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo } },
        { text: 'Disciplina', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo } },
        { text: '1º Bim',     options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
        { text: '2º Bim',     options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
        { text: '3º Bim',     options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
        { text: '4º Bim',     options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
        { text: 'Total %',    options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } },
      ],
      ...combos.map((c, idx) => {
        const bims = calcDigitalBimestres(c.regs);
        const total = calcDigital(c.regs);
        const bg = idx % 2 === 0 ? PPTX_CORES.branco : 'F3E8FF';
        const bimCell = (b) => {
          if (!b) return { text: '—', options: { fill: bg, align: 'center', color: PPTX_CORES.cinza } };
          const cor = b.pct >= 80 ? PPTX_CORES.verde : b.pct >= 50 ? PPTX_CORES.amber : PPTX_CORES.verm;
          return { text: `${b.pct.toFixed(0)}%\n(${b.concluido}/${b.previsto})`, options: { fill: bg, align: 'center', color: cor, bold: true, fontSize: 9 } };
        };
        const totalCor = total ? (total.pct >= 80 ? PPTX_CORES.verde : total.pct >= 50 ? PPTX_CORES.amber : PPTX_CORES.verm) : PPTX_CORES.cinza;
        return [
          { text: c.turma, options: { fill: bg, bold: true, align: 'center' } },
          { text: c.disc,  options: { fill: bg } },
          bimCell(bims[0]), bimCell(bims[1]), bimCell(bims[2]), bimCell(bims[3]),
          { text: total ? `${total.pct.toFixed(0)}%` : '—', options: { fill: bg, align: 'center', color: totalCor, bold: true } },
        ];
      }),
    ];

    slideD2.addTable(digTabRows, {
      x: 0.3, y: 0.75, w: 9.4, h: Math.min(3.5, 0.38 * (combos.length + 1)),
      fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc,
      border: { pt: 0.5, color: 'E9D5FF' },
      rowH: 0.38,
      colW: [0.9, 2.3, 1.1, 1.1, 1.1, 1.1, 1.0],
    });
  }
}

// ── Função principal ────────────────────────────────────────

function gerarPPTX() {
  if (!APP.notas.length) {
    alert('Importe dados antes de gerar a apresentação.');
    return;
  }

  loader(true, 'Gerando apresentação PPTX...');

  setTimeout(() => {
    try {
      const turmaSel = document.getElementById('pptx-sel-turma')?.value || '';
      const discSel  = document.getElementById('pptx-sel-disc')?.value  || '';
      const modoDetalhado = !!(turmaSel || discSel);
      const data = new Date().toLocaleDateString('pt-BR');

      const prs = new PptxGenJS();
      prs.layout = 'LAYOUT_WIDE';
      prs.author  = CFG.escola;
      prs.company = CFG.escola;
      prs.subject = 'Conselho de Classe';

      if (!modoDetalhado) {
        // ══ RELATÓRIO GERAL ══════════════════════════════════
        pptxCapa(
          prs,
          'Relatório Geral — Conselho de Classe',
          `${CFG.escola} · ${CFG.etapa}`,
          `Gerado em ${data} · Todas as turmas`
        );

        const turmas = unique(APP.notas.map(r => r.TURMA));
        const alunos = unique(APP.notas.map(r => r.ESTUDANTE));
        let ap = 0, rp = 0, alF = 0;
        alunos.forEach(a => {
          const regs = APP.notas.filter(r => r.ESTUDANTE === a);
          const ds = regs.map(r => dadosReg(r));
          if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++;
          if (ds.some(d => d.alFreq)) alF++;
        });
        const medias = APP.notas.map(r => calcMedia(r)).filter(m => m !== null);
        const mg = medias.length ? +(medias.reduce((a, b) => a + b, 0) / medias.length) : 0;

        const resumoTurmas = turmas.map(t => {
          const al = unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE));
          let ta = 0, tr = 0, taf = 0;
          al.forEach(a => {
            const ds = APP.notas.filter(r => r.ESTUDANTE === a && r.TURMA === t).map(r => dadosReg(r));
            if (ds.some(d => d.sit === 'Reprovado')) tr++; else ta++;
            if (ds.some(d => d.alFreq)) taf++;
          });
          const ms2 = APP.notas.filter(r => r.TURMA === t).map(r => calcMedia(r)).filter(m => m !== null);
          const tmg = ms2.length ? +(ms2.reduce((a, b) => a + b, 0) / ms2.length).toFixed(1) : 0;
          return { turma: t, alunos: al.length, aprov: ta, reprov: tr, alFreq: taf, mg: +tmg };
        });

        pptxSlideResumoEscola(prs, { turmas: resumoTurmas, totalAlunos: alunos.length, aprov: ap, reprov: rp, alFreqs: alF, mg });

        turmas.forEach(t => pptxSlideResumoPorTurma(prs, t));

        prs.writeFile({ fileName: `Conselho_Classe_Geral_${data.replace(/\//g, '-')}.pptx` });

      } else {
        // ══ RELATÓRIO DETALHADO ═══════════════════════════════
        const labelTurma = turmaSel || 'Todas as turmas';
        const labelDisc  = discSel  || 'Todas as disciplinas';
        const subtitulo  = [turmaSel && `Turma ${turmaSel}`, discSel].filter(Boolean).join(' · ');

        pptxCapa(
          prs,
          'Relatório Detalhado — Conselho de Classe',
          `${CFG.escola} · ${subtitulo}`,
          `Gerado em ${data} · ${CFG.etapa}`
        );

        pptxSlidesDetalhado(prs, turmaSel, discSel);

        const nomeSafe = subtitulo.replace(/[^a-zA-Z0-9À-ÿ\s]/g, '').replace(/\s+/g, '_');
        prs.writeFile({ fileName: `Conselho_Classe_${nomeSafe}_${data.replace(/\//g, '-')}.pptx` });
      }

    } catch (err) {
      alert('Erro ao gerar PPTX: ' + err.message);
      console.error(err);
    }
    loader(false);
  }, 80);
}

// ── UI da seção PPTX ────────────────────────────────────────

function popSelectsPPTX() {
  if (!APP.notas.length) return;
  const turmas = unique(APP.notas.map(r => r.TURMA));
  const discs  = unique(APP.notas.map(r => r.DISCIPLINA));

  const selT = document.getElementById('pptx-sel-turma');
  const selD = document.getElementById('pptx-sel-disc');
  if (selT) selT.innerHTML = '<option value="">Todas as turmas (Relatório Geral)</option>' +
    turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  if (selD) selD.innerHTML = '<option value="">Todas as disciplinas</option>' +
    discs.map(d => `<option value="${d}">${d}</option>`).join('');
  atualizarContextoPPTX();
}

function atualizarContextoPPTX() {
  if (!APP.notas.length) return;
  const turmaSel = document.getElementById('pptx-sel-turma')?.value || '';
  const discSel  = document.getElementById('pptx-sel-disc')?.value  || '';
  const modoDetalhado = !!(turmaSel || discSel);
  const turmas = unique(APP.notas.map(r => r.TURMA));
  const alunos = unique(APP.notas.filter(r =>
    (!turmaSel || r.TURMA === turmaSel) &&
    (!discSel  || r.DISCIPLINA === discSel)
  ).map(r => r.ESTUDANTE));

  const prevTitulo = document.getElementById('pptx-prev-titulo');
  const prevDesc   = document.getElementById('pptx-prev-desc');
  const prevSlides = document.getElementById('pptx-prev-slides');

  if (!modoDetalhado) {
    if (prevTitulo) prevTitulo.textContent = 'Relatório Geral da Escola';
    if (prevDesc)   prevDesc.textContent = `${turmas.length} turmas · ${unique(APP.notas.map(r => r.ESTUDANTE)).length} alunos`;
    if (prevSlides) prevSlides.innerHTML = `
      <div class="pptx-slide-item"><span class="pptx-slide-num">1</span> Capa</div>
      <div class="pptx-slide-item"><span class="pptx-slide-num">2</span> Painel consolidado</div>
      ${turmas.map((t, i) => `<div class="pptx-slide-item"><span class="pptx-slide-num">${i + 3}</span> Turma ${t}</div>`).join('')}
    `;
  } else {
    const label = [turmaSel && `Turma ${turmaSel}`, discSel].filter(Boolean).join(' · ');
    if (prevTitulo) prevTitulo.textContent = `Relatório Detalhado — ${label}`;
    if (prevDesc)   prevDesc.textContent = `${alunos.length} alunos · dados filtrados`;
    const regs = APP.notas.filter(r => (!turmaSel || r.TURMA === turmaSel) && (!discSel || r.DISCIPLINA === discSel));
    const pags = Math.ceil(regs.length / 14);
    const temDig = APP.digital.some(r => (!turmaSel || r.TURMA === turmaSel) && (!discSel || r.DISCIPLINA === discSel));
    if (prevSlides) prevSlides.innerHTML = `
      <div class="pptx-slide-item"><span class="pptx-slide-num">1</span> Capa</div>
      <div class="pptx-slide-item"><span class="pptx-slide-num">2</span> Resumo e indicadores</div>
      ${Array.from({length: pags}, (_, i) => `<div class="pptx-slide-item"><span class="pptx-slide-num">${i + 3}</span> Lista de alunos${pags > 1 ? ` (${i+1}/${pags})` : ''}</div>`).join('')}
      ${temDig ? `<div class="pptx-slide-item"><span class="pptx-slide-num">${3 + pags}</span> Material Digital</div>` : ''}
    `;
  }
}

// ══════════════════════════════════════════════════════════════
//  RADAR PREDITIVO — Centro de Inteligência Pedagógica
//  Regras anti-punição ao professor: alertas de discrepância
//  para investigação, não veredictos.
// ══════════════════════════════════════════════════════════════

function renderRadar() {
  if (!APP.notas.length) {
    document.getElementById('mg-radar').innerHTML =
      `<div class="mc" style="grid-column:1/-1;border-left:4px solid #c2410c;">
        <div class="ml">Sem dados</div>
        <div class="mv" style="font-size:14px;font-weight:500;color:#c2410c;">Importe dados para ativar o Radar Preditivo</div>
      </div>`;
    return;
  }

  const turmaFiltro = document.getElementById('radar-sel-turma')?.value || '';
  const discFiltro  = document.getElementById('radar-sel-disc')?.value  || '';

  // Registros filtrados
  const regs = APP.notas.filter(r =>
    (!turmaFiltro || r.TURMA === turmaFiltro) &&
    (!discFiltro  || r.DISCIPLINA === discFiltro)
  );
  if (!regs.length) {
    document.getElementById('mg-radar').innerHTML =
      `<div class="mc" style="grid-column:1/-1;border-left:4px solid var(--cinza);">
        <div class="ml">Nenhum registro</div>
        <div class="mv c-rx" style="font-size:14px;">Ajuste os filtros para visualizar dados</div>
      </div>`;
    return;
  }

  const turmas = unique(regs.map(r => r.TURMA));
  const discs  = unique(regs.map(r => r.DISCIPLINA));
  const alunos = unique(regs.map(r => r.ESTUDANTE));

  // Métricas globais do filtro
  let totalRisco = 0, totalColapso = 0, totalDesistencia = 0;
  alunos.forEach(a => {
    const aRegs = regs.filter(r => r.ESTUDANTE === a);
    const aDados = aRegs.map(r => dadosReg(r));
    const reprovCount = aDados.filter(d => d.sit === 'Reprovado').length;
    const alFreqCount = aDados.filter(d => d.alFreq).length;
    const medias = aDados.map(d => d.media).filter(m => m !== null);
    const mediaGeral = medias.length ? +(medias.reduce((a,b)=>a+b,0)/medias.length).toFixed(2) : null;
    if (reprovCount > 0) totalRisco++;
    if (reprovCount >= Math.ceil(aDados.length * 0.5) && alFreqCount > 0) totalColapso++;
    if (mediaGeral !== null && mediaGeral >= CFG.mediaAprov && alFreqCount > 0) totalDesistencia++;
  });

  const pctRisco = alunos.length ? Math.round(totalRisco / alunos.length * 100) : 0;
  document.getElementById('mg-radar').innerHTML = `
    <div class="mc bl"><div class="ml">Alunos no filtro</div><div class="mv c-bl">${alunos.length}</div><div class="ms">${turmas.length} turma(s) · ${discs.length} disciplina(s)</div></div>
    <div class="mc vm"><div class="ml">Em risco de reprovação</div><div class="mv c-vm">${totalRisco}</div><div class="ms">${pctRisco}% do filtro</div></div>
    <div class="mc am"><div class="ml">Desistência silenciosa</div><div class="mv c-am">${totalDesistencia}</div><div class="ms">nota ok, faltas crescentes</div></div>
    <div class="mc rx"><div class="ml">Colapso global</div><div class="mv c-rx">${totalColapso}</div><div class="ms">reprov. maioria + faltas altas</div></div>`;

  // ── 1. Alertas Discrepância Digital vs Notas ─────────────
  const bodyDisc = document.getElementById('radar-discrepancia-body');
  if (!APP.digital.length) {
    bodyDisc.innerHTML = `<p style="color:var(--cinza);font-size:13px;"><i class="bi bi-info-circle"></i> Importe uma planilha com a aba <strong>DIGITAL</strong> para ativar esta análise cruzada.</p>`;
  } else {
    let htmlDisc = '';
    // Para cada combinação turma × disciplina no filtro
    const combos = [];
    turmas.forEach(t => discs.forEach(d => {
      if (regs.some(r => r.TURMA === t && r.DISCIPLINA === d)) {
        combos.push({ turma: t, disc: d });
      }
    }));

    combos.forEach(({ turma, disc }) => {
      const digRegs = APP.digital.filter(r => r.TURMA === turma && r.DISCIPLINA === disc);
      if (!digRegs.length) return;
      const calc = calcDigital(digRegs);
      if (!calc) return;

      const notasCombo = regs.filter(r => r.TURMA === turma && r.DISCIPLINA === disc);
      const mediasCombo = notasCombo.map(r => calcMedia(r)).filter(m => m !== null);
      if (!mediasCombo.length) return;
      const mediaMedia = +(mediasCombo.reduce((a,b)=>a+b,0)/mediasCombo.length).toFixed(2);

      // Regra: digital > 70% e média < 5.0 → discrepância
      if (calc.pct > 70 && mediaMedia < CFG.mediaAprov) {
        const alunosBaixoRend = notasCombo
          .filter(r => { const m = calcMedia(r); return m !== null && m < CFG.mediaAprov; })
          .map(r => r.ESTUDANTE);
        htmlDisc += `
          <div class="radar-alerta radar-alerta-disc">
            <div class="radar-alerta-icon">⚠️</div>
            <div class="radar-alerta-body">
              <div class="radar-alerta-titulo">Discrepância em ${disc} — Turma ${turma}</div>
              <div class="radar-alerta-desc">
                Alto volume de material digital aplicado (<strong>${calc.pct.toFixed(0)}%</strong> de cumprimento),
                mas baixo rendimento da turma (média <strong>${mediaMedia.toFixed(1).replace('.',',')}</strong>).
                <br><strong>Sugestão:</strong> Investigar engajamento dos alunos nas avaliações (possíveis "chutes" ou desinteresse).
                Esta discrepância <em>não implica falha do professor</em> — o material foi aplicado.
              </div>
              <div class="radar-alerta-alunos">
                <span class="radar-alunos-label">Alunos com baixo rendimento nesta métrica:</span>
                <div class="radar-alunos-lista">
                  ${alunosBaixoRend.map(n => `<span class="radar-aluno-tag radar-tag-amber">${n}</span>`).join('')}
                </div>
              </div>
            </div>
          </div>`;
      }

      // Regra adicional: digital < 30% e média >= 5.0 → positivo sem muito digital
      if (calc.pct < 30 && mediaMedia >= CFG.mediaAprov) {
        htmlDisc += `
          <div class="radar-alerta radar-alerta-info">
            <div class="radar-alerta-icon">ℹ️</div>
            <div class="radar-alerta-body">
              <div class="radar-alerta-titulo">Ponto de curiosidade em ${disc} — Turma ${turma}</div>
              <div class="radar-alerta-desc">
                Baixo uso de material digital (<strong>${calc.pct.toFixed(0)}%</strong>) com bom rendimento (média <strong>${mediaMedia.toFixed(1).replace('.',',')}</strong>).
                <strong>Sugestão:</strong> Verificar se há metodologias complementares em uso ou se o material digital ainda será ampliado.
              </div>
            </div>
          </div>`;
      }
    });

    bodyDisc.innerHTML = htmlDisc ||
      `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div>
       <div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhuma discrepância digital detectada no filtro atual</div>
       <div class="radar-alerta-desc">Os dados de material digital e notas estão alinhados para as turmas/disciplinas selecionadas.</div></div></div>`;
  }

  // ── 2. Alerta de volume de alunos em risco por turma/disc ─
  const bodyRisco = document.getElementById('radar-risco-turma-body');
  let htmlRisco = '';
  const combosRisco = [];
  turmas.forEach(t => discs.forEach(d => {
    if (regs.some(r => r.TURMA === t && r.DISCIPLINA === d)) combosRisco.push({ turma: t, disc: d });
  }));

  combosRisco.forEach(({ turma, disc }) => {
    const notasCombo = regs.filter(r => r.TURMA === turma && r.DISCIPLINA === disc);
    if (!notasCombo.length) return;
    const dados = notasCombo.map(r => dadosReg(r));
    const emRisco = dados.filter(d => d.sit === 'Reprovado');
    const pct = Math.round(emRisco.length / dados.length * 100);
    if (pct > 30) {
      htmlRisco += `
        <div class="radar-alerta radar-alerta-risco">
          <div class="radar-alerta-icon">📊</div>
          <div class="radar-alerta-body">
            <div class="radar-alerta-titulo">Ponto de Atenção em ${disc} — Turma ${turma}</div>
            <div class="radar-alerta-desc">
              <strong>${emRisco.length} alunos</strong> (${pct}% da turma nesta disciplina) estão com projeção de reprovação.
              <strong>Sugestão:</strong> Alinhar plano de recuperação com o professor. Esta situação reflete o desempenho dos estudantes, não a metodologia.
            </div>
            <div class="radar-alerta-alunos">
              <span class="radar-alunos-label">Estudantes em risco:</span>
              <div class="radar-alunos-lista">
                ${emRisco.map(d => `<span class="radar-aluno-tag radar-tag-verm">${d.n.ESTUDANTE}</span>`).join('')}
              </div>
            </div>
          </div>
        </div>`;
    }
  });

  bodyRisco.innerHTML = htmlRisco ||
    `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div>
     <div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhuma turma/disciplina com mais de 30% em risco</div>
     <div class="radar-alerta-desc">Todas as combinações turma × disciplina no filtro estão dentro do limite de atenção.</div></div></div>`;

  // ── 3. Perfis de Risco Individual ────────────────────────
  const desistentes = [], localizados = [], colapsos = [];

  alunos.forEach(a => {
    const aRegs = regs.filter(r => r.ESTUDANTE === a);
    const aDados = aRegs.map(r => dadosReg(r));
    const turmaAluno = aRegs[0]?.TURMA || '';
    const reprovCount = aDados.filter(d => d.sit === 'Reprovado').length;
    const alFreqCount = aDados.filter(d => d.alFreq).length;
    const medias = aDados.map(d => d.media).filter(m => m !== null);
    const mediaGeral = medias.length ? +(medias.reduce((x,y)=>x+y,0)/medias.length).toFixed(2) : null;
    const totalDiscs = aDados.length;

    // Colapso global: reprovando em >= 50% das disciplinas E faltas altas
    if (reprovCount >= Math.ceil(totalDiscs * 0.5) && alFreqCount > 0) {
      colapsos.push({ nome: a, turma: turmaAluno, reprov: reprovCount, totalDiscs, alFreq: alFreqCount });
    }
    // Desistência silenciosa: nota >= 5 geral mas faltas crescentes
    else if (mediaGeral !== null && mediaGeral >= CFG.mediaAprov && alFreqCount > 0) {
      desistentes.push({ nome: a, turma: turmaAluno, media: mediaGeral, alFreq: alFreqCount });
    }
    // Dificuldade localizada: reprovando em 1 ou 2 disciplinas
    else if (reprovCount === 1 || reprovCount === 2) {
      const discsReprov = aRegs
        .filter(r => { const d = dadosReg(r); return d.sit === 'Reprovado'; })
        .map(r => r.DISCIPLINA);
      localizados.push({ nome: a, turma: turmaAluno, discs: discsReprov });
    }
  });

  function perfilItem(item, tipo) {
    if (tipo === 'desistencia') {
      return `<div class="radar-perfil-item">
        <strong>${item.nome}</strong>
        <span class="radar-perfil-detalhe">${item.turma} · Média ${item.media.toFixed(1).replace('.',',')} · ${item.alFreq} disc. com alerta freq.</span>
      </div>`;
    }
    if (tipo === 'localizada') {
      return `<div class="radar-perfil-item">
        <strong>${item.nome}</strong>
        <span class="radar-perfil-detalhe">${item.turma} · Reprovando em: ${item.discs.join(', ')}</span>
      </div>`;
    }
    if (tipo === 'colapso') {
      return `<div class="radar-perfil-item">
        <strong>${item.nome}</strong>
        <span class="radar-perfil-detalhe">${item.turma} · ${item.reprov}/${item.totalDiscs} disc. reprovadas · ${item.alFreq} com alerta freq.</span>
      </div>`;
    }
    return '';
  }

  document.getElementById('radar-lista-desistencia').innerHTML = desistentes.length
    ? desistentes.map(i => perfilItem(i, 'desistencia')).join('')
    : `<div class="radar-perfil-vazio">Nenhum aluno neste perfil no filtro selecionado.</div>`;

  document.getElementById('radar-lista-localizada').innerHTML = localizados.length
    ? localizados.map(i => perfilItem(i, 'localizada')).join('')
    : `<div class="radar-perfil-vazio">Nenhum aluno neste perfil no filtro selecionado.</div>`;

  document.getElementById('radar-lista-colapso').innerHTML = colapsos.length
    ? colapsos.map(i => perfilItem(i, 'colapso')).join('')
    : `<div class="radar-perfil-vazio">Nenhum aluno neste perfil no filtro selecionado.</div>`;

  // ── 4. Índice de Viabilidade de Recuperação ──────────────
  const bodyViab = document.getElementById('radar-viabilidade-body');

  // Para cada aluno + disciplina no filtro, calcular nota necessária
  const viabRows = [];

  regs.forEach(r => {
    const { media, sit, parc } = dadosReg(r);
    if (media === null) return;
    if (sit !== 'Reprovado' && !parc) return; // só parciais ou reprovados

    // Quantos bimestres já foram lançados
    const bims = [r.b1, r.b2, r.b3, r.b4];
    const lancados = bims.filter(v => v !== null).length;
    const restantes = 4 - lancados;

    if (restantes === 0) {
      // Sem bimestres restantes — reprovado definitivo
      if (sit === 'Reprovado') {
        viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: null, impossivel: true });
      }
      return;
    }

    // Calcular nota necessária por bimestre para atingir média 5.0
    // (soma atual + restantes × X) / 4 >= 5.0
    const somaAtual = bims.filter(v => v !== null).reduce((a, b) => a + b, 0);
    const somaNecess = CFG.mediaAprov * 4;
    const somaRestante = somaNecess - somaAtual;
    const notaPorBim = somaRestante / restantes;

    if (notaPorBim > 10) {
      viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: notaPorBim, restantes, impossivel: true });
    } else if (notaPorBim > 8.5) {
      viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: notaPorBim, restantes, impossivel: false, dificil: true });
    } else if (notaPorBim > 0) {
      viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: notaPorBim, restantes, impossivel: false, dificil: false });
    }
  });

  if (!viabRows.length) {
    bodyViab.innerHTML = `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div>
      <div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhum aluno em situação de risco de recuperação no filtro atual</div></div></div>`;
    return;
  }

  // Agrupar em 3 níveis
  const impossiveis = viabRows.filter(v => v.impossivel);
  const dificeis = viabRows.filter(v => !v.impossivel && v.dificil);
  const possiveis = viabRows.filter(v => !v.impossivel && !v.dificil);

  function viabRow(v) {
    const notaStr = v.notaNecessaria === null ? 'Reprovado (sem bimestres restantes)' :
      v.notaNecessaria > 10 ? `${v.notaNecessaria.toFixed(1).replace('.',',')} (impossível — acima de 10,0)` :
      `${v.notaNecessaria.toFixed(1).replace('.',',')} em cada bimestre restante (${v.restantes} restante${v.restantes > 1 ? 's' : ''})`;
    return `<tr>
      <td><strong>${v.nome}</strong></td>
      <td><span class="badge b-info">${v.turma}</span></td>
      <td style="font-size:12px;">${v.disc}</td>
      <td>${fmtMedia(v.mediaAtual, true)}</td>
      <td style="font-size:12px;${v.impossivel ? 'color:var(--verm);font-weight:600;' : v.dificil ? 'color:var(--amber);font-weight:600;' : 'color:var(--verde);'}">${notaStr}</td>
      <td>${v.impossivel
        ? `<span class="badge b-re"><i class="bi bi-x-circle-fill"></i> Intervenção externa</span>`
        : v.dificil
          ? `<span class="badge b-al"><i class="bi bi-exclamation-triangle-fill"></i> Recuperação difícil</span>`
          : `<span class="badge b-ap"><i class="bi bi-check-circle-fill"></i> Recuperação viável</span>`
      }</td>
    </tr>`;
  }

  const allRows = [...impossiveis, ...dificeis, ...possiveis];

  bodyViab.innerHTML = `
    <div class="ib" style="margin-bottom:14px;">
      <i class="bi bi-info-circle-fill"></i>
      <span>Cálculo baseado na nota mínima necessária em cada bimestre restante para atingir média ${CFG.mediaAprov.toFixed(1).replace('.',',')}. "Intervenção externa" significa que a média já não pode ser atingida matematicamente nas avaliações regulares.</span>
    </div>
    ${impossiveis.length ? `<div style="font-size:12px;font-weight:700;color:var(--verm);margin-bottom:6px;"><i class="bi bi-x-circle-fill"></i> Recuperação Improvável — Necessita Intervenção Externa (${impossiveis.length})</div>` : ''}
    <div class="tw"><table>
      <thead><tr><th>Aluno</th><th>Turma</th><th>Disciplina</th><th>Média atual</th><th>Nota necessária</th><th>Viabilidade</th></tr></thead>
      <tbody>${allRows.map(viabRow).join('')}</tbody>
    </table></div>`;
}

// ── DEMO ───────────────────────────────────────────────────
function carregarDemo(){
  const nomes=[
    'ALEXIA RAISSA ALMEIDA ALVES','ALYANDRO ARIEL SANTOS FREITAS','CRISTOFFER MIGUEL FERREIRA LIMA',
    'IRANI CAIRAC PRESTES NETO','JOAO LUCAS BORGES DE OLIVEIRA','LEANDRO HENRIQUE DE ALMEIDA',
    'LUIS GUSTAVO DE OLIVEIRA SANTOS','MARIA FERNANDA COSTA SILVA','PEDRO HENRIQUE RAMOS JUNIOR',
    'ANA CAROLINA MENDES SOUSA','GABRIEL TORRES NASCIMENTO','BEATRIZ LIMA CAVALCANTE',
    'LUCAS AUGUSTO FERREIRA DIAS','JULIA APARECIDA MOREIRA COSTA','THIAGO CAMARGO BORGES',
    'LARISSA BRITO CARVALHO','VINICIUS MARTINS PRADO','CAMILA RODRIGUES PEREIRA',
    'FELIPE SOUZA ALBUQUERQUE','NATHALIA GOMES RIBEIRO'
  ];
  const turmas=['2ªA','2ªC','3ªA'];
  const discs=['MATEMÁTICA','PROGRAMAÇÃO','ROBÓTICA'];
  APP.notas=[];APP.faltas=[];APP.digital=[];

  turmas.forEach(t=>{
    const al=nomes.slice(0, t==='3ªA'?18:20);
    discs.forEach(d=>{
      al.forEach(a=>{
        const rnd=Math.random();
        const b1=+(rnd<.1?1.5+Math.random()*2:rnd<.25?3+Math.random()*1.5:5+Math.random()*5).toFixed(1);
        const b2=+(Math.min(10,Math.max(0,b1+(Math.random()*2-0.8)))).toFixed(1);
        const b3=+(Math.min(10,Math.max(0,b2+(Math.random()*2-0.7)))).toFixed(1);
        const b4=Math.random()>.3?+(Math.min(10,Math.max(0,b3+(Math.random()*2-0.5)))).toFixed(1):null;
        APP.notas.push({ESTUDANTE:a,TURMA:t,DISCIPLINA:d,b1,b2,b3,b4});
        const a1=38+Math.floor(Math.random()*4);
        const a2=39+Math.floor(Math.random()*4);
        const a3=38+Math.floor(Math.random()*5);
        const a4=b4!==null?39+Math.floor(Math.random()*4):null;
        const faltas=Math.random()<.15?Math.floor(Math.random()*15)+5:Math.floor(Math.random()*5);
        APP.faltas.push({ESTUDANTE:a,TURMA:t,DISCIPLINA:d,
          a1,f1:Math.floor(Math.random()*4),
          a2,f2:Math.floor(faltas*.4),
          a3,f3:Math.floor(faltas*.4),
          a4,f4:a4?Math.floor(faltas*.2):null
        });
      });
      // Dados de material digital por turma/disciplina (independente dos alunos)
      const prev=[12,12,12,12];
      const pctConc=[1,0.85,0.6,0];
      APP.digital.push({
        TURMA:t, DISCIPLINA:d,
        prev1:prev[0], conc1:Math.round(prev[0]*pctConc[0]*(0.9+Math.random()*0.2)),
        prev2:prev[1], conc2:Math.round(prev[1]*(pctConc[1]+Math.random()*0.15)),
        prev3:prev[2], conc3:Math.round(prev[2]*(pctConc[2]+Math.random()*0.2)),
        prev4:prev[3], conc4:null,
      });
    });
  });

  document.getElementById('status-up').innerHTML=
    `<span style="color:var(--verde);">✓ Dados de exemplo carregados: ${APP.notas.length} registros em ${turmas.length} turmas · ${APP.digital.length} registros de material digital.</span>`;
  buildNav();
  showSection('geral');
}

// ── MODELO EXCEL ───────────────────────────────────────────
function baixarModelo(){
  const wb = XLSX.utils.book_new();

  // Aba Notas
  const notasData = [
    ['ESTUDANTE','TURMA','DISCIPLINA','NOTA 1º BIMESTRE','NOTA 2º BIMESTRE','NOTA 3º BIMESTRE','NOTA 4º BIMESTRE'],
    ['NOME DO ALUNO','2ªA','MATEMÁTICA',7,8,6,7],
    ['OUTRO ALUNO','2ªA','MATEMÁTICA',5,4,'',''],
    ['NOME DO ALUNO','2ªA','ROBÓTICA',8,9,7,8],
    ['OUTRO ALUNO','2ªA','ROBÓTICA',6,5,6,''],
  ];
  const wsNotas = XLSX.utils.aoa_to_sheet(notasData);
  wsNotas['!cols']=[{wch:35},{wch:8},{wch:20},{wch:18},{wch:18},{wch:18},{wch:18}];
  XLSX.utils.book_append_sheet(wb, wsNotas, 'Notas');

  // Aba FALTAS
  const faltasData = [
    ['ESTUDANTE','TURMA','DISCIPLINA','AULAS 1º BIM','FALTAS 1º BIM','AULAS 2º BIM','FALTAS 2º BIM','AULAS 3º BIM','FALTAS 3º BIM','AULAS 4º BIM','FALTAS 4º BIM'],
    ['NOME DO ALUNO','2ªA','MATEMÁTICA',38,0,40,2,39,5,39,3],
    ['OUTRO ALUNO','2ªA','MATEMÁTICA',38,4,40,14,'','','',''],
    ['NOME DO ALUNO','2ªA','ROBÓTICA',20,0,22,1,20,2,22,1],
    ['OUTRO ALUNO','2ªA','ROBÓTICA',20,3,22,8,'','','',''],
  ];
  const wsFaltas = XLSX.utils.aoa_to_sheet(faltasData);
  wsFaltas['!cols']=[{wch:35},{wch:8},{wch:20},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];
  XLSX.utils.book_append_sheet(wb, wsFaltas, 'FALTAS');

  // Aba DIGITAL (nova)
  const digitalData = [
    ['TURMA','DISCIPLINA','PREVISTO 1º BIM','CONCLUIDO 1º BIM','PREVISTO 2º BIM','CONCLUIDO 2º BIM','PREVISTO 3º BIM','CONCLUIDO 3º BIM','PREVISTO 4º BIM','CONCLUIDO 4º BIM'],
    ['2ªA','MATEMÁTICA',10,10,10,8,10,5,10,''],
    ['2ªA','ROBÓTICA',8,8,8,6,8,4,8,''],
    ['2ªB','MATEMÁTICA',10,10,10,10,10,7,10,''],
    ['2ªB','ROBÓTICA',8,8,8,7,8,5,8,''],
  ];
  const wsDigital = XLSX.utils.aoa_to_sheet(digitalData);
  wsDigital['!cols']=[{wch:8},{wch:20},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb, wsDigital, 'DIGITAL');

  XLSX.writeFile(wb, 'modelo_conselho_classe.xlsx');
}

// ── LIMPAR ─────────────────────────────────────────────────
function limpar(){
  if(!confirm('Limpar todos os dados?')) return;
  APP.notas=[];APP.faltas=[];APP.digital=[];APP.turmaSel=null;APP.discSel=null;
  Object.values(APP.charts).forEach(c=>c.destroy());APP.charts={};
  document.getElementById('nav-turmas').innerHTML='';
  document.getElementById('nav-discs').innerHTML='';
  document.getElementById('status-up').innerHTML='';
  document.getElementById('inp-p').value='';
  // Reset accordion counts and state
  ['acc-turmas','acc-discs'].forEach(id => {
    const acc = document.getElementById(id);
    if (!acc) return;
    const toggle = acc.querySelector('.sb-acc-toggle');
    const body = acc.querySelector('.sb-acc-body');
    if (toggle) { toggle.classList.remove('open'); toggle.setAttribute('aria-expanded','false'); }
    if (body) body.classList.remove('open');
  });
  const countT = document.getElementById('acc-turmas-count');
  const countD = document.getElementById('acc-discs-count');
  if (countT) { countT.textContent = ''; countT.classList.remove('visible'); }
  if (countD) { countD.textContent = ''; countD.classList.remove('visible'); }
  showSection('upload');
}

// ── SIDEBAR MOBILE ──────────────────────────────────────────
function abrirSidebar() {
  document.getElementById('sidebar').classList.add('open');
  document.getElementById('sb-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}
function fecharSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sb-overlay').classList.remove('open');
  document.body.style.overflow = '';
}
// Fechar sidebar ao clicar em um link (mobile)
document.addEventListener('click', e => {
  if (window.innerWidth <= 768 && e.target.closest('#sidebar .nav-link')) {
    fecharSidebar();
  }
});

// ── LGPD ────────────────────────────────────────────────────
function aceitarCookies() {
  localStorage.setItem('lgpd_consent', 'accepted');
  document.getElementById('lgpd-banner').classList.add('hidden');
}
function recusarCookies() {
  localStorage.setItem('lgpd_consent', 'essential');
  document.getElementById('lgpd-banner').classList.add('hidden');
}
function abrirPolitica() {
  document.getElementById('lgpd-modal').classList.add('open');
}
function fecharPolitica(e) {
  if (e.target === document.getElementById('lgpd-modal'))
    document.getElementById('lgpd-modal').classList.remove('open');
}

// ── INIT ───────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('loader').style.display = 'none';
  // Verificar consentimento LGPD
  const consent = localStorage.getItem('lgpd_consent');
  if (consent) document.getElementById('lgpd-banner').classList.add('hidden');
});
