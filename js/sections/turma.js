/**
 * sections/turma.js — Drill-down por turma + modal de detalhes do aluno.
 */

import { getTurma, getTurmaAnterior, getTurmaChaves, hasTurmas, hasTurmasAnterior, getConfig } from '../store.js';
import { calcSituacao, calcMedia, calcFreqMedia, calcFreq, resolverNota } from '../calc.js';
import {
  criarBarHorizontal,
  criarBarChart,
  pluginBarPct,
  dataset,
  COLORS,
  STATUS_COLORS,
} from '../charts.js';
import { showModal, closeModal } from '../ui.js';

const MODAL_ID = 'modal-aluno-overlay';

// ---------------------------------------------------------------------------
// Init: populate selector + event
// ---------------------------------------------------------------------------

export function initTurma() {
  const select = document.getElementById('select-turma');
  if (!select) return;

  select.addEventListener('change', () => renderTurma());

  // Atualizar opções quando dados mudam
  window.addEventListener('cc:data-changed', () => {
    populateSelector(select);
    renderTurma();
  });
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderTurma() {
  if (!hasTurmas()) {
    clearTurma();
    return;
  }

  const select = document.getElementById('select-turma');
  populateSelector(select);
  const chave = select?.value || getTurmaChaves()[0];
  if (select && select.value !== chave) select.value = chave;

  const turma = getTurma(chave);
  if (!turma) { clearTurma(); return; }

  const config = getConfig();

  // -----------------------------------------------------------------------
  // Cards
  // -----------------------------------------------------------------------
  const alunos = turma.alunos;
  let aprovados = 0;
  let alertaFreq = 0;
  let somaMedia = 0;
  let countMedia = 0;

  for (const aluno of alunos) {
    const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
    if (sit === 'Aprovado') aprovados++;
    if (nivelFreq === 'alerta') alertaFreq++;

    const media = calcMedia(aluno, turma.disciplinas, config);
    if (media !== null) { somaMedia += media; countMedia++; }
  }

  const mediaTurma = countMedia > 0 ? somaMedia / countMedia : 0;

  document.getElementById('turma-alunos').textContent = alunos.length;
  document.getElementById('turma-media').textContent = countMedia > 0 ? mediaTurma.toFixed(1) : '—';
  document.getElementById('turma-aprovados').textContent = aprovados;
  document.getElementById('turma-alerta-freq').textContent = alertaFreq;

  // -----------------------------------------------------------------------
  // Tabela de alunos
  // -----------------------------------------------------------------------
  const turmaAnterior = getTurmaAnterior(chave);
  const tbody = document.getElementById('turma-tabela-body');
  const thead = document.querySelector('#section-turma .table thead tr');

  // Atualizar cabeçalho da tabela se houver bimestre anterior
  if (thead) {
    if (turmaAnterior) {
      thead.innerHTML = `
        <th>Nome</th>
        <th>Situação</th>
        <th class="table-numeric">Média</th>
        <th class="table-numeric">Freq.</th>
        <th class="table-numeric">Média Ant.</th>
        <th class="table-numeric">Freq. Ant.</th>
        <th class="table-center">Detalhes</th>`;
    } else {
      thead.innerHTML = `
        <th>Nome</th>
        <th>Situação</th>
        <th class="table-numeric">Média</th>
        <th class="table-numeric">Frequência</th>
        <th class="table-center">Detalhes</th>`;
    }
  }

  if (tbody) {
    tbody.innerHTML = alunos
      .map((aluno) => {
        const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
        const media = calcMedia(aluno, turma.disciplinas, config);
        const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

        let sitBadge = '';
        if (sit === 'Aprovado') {
          if (nivelFreq === 'alerta') sitBadge = '<span class="badge badge-warning">Aprovado ⚠ freq.</span>';
          else if (nivelFreq === 'critico') sitBadge = '<span class="badge badge-critical">Aprovado ⚠ retenção</span>';
          else sitBadge = '<span class="badge badge-success">Aprovado</span>';
        } else {
          sitBadge = '<span class="badge badge-critical">Reprovado</span>';
        }

        // Comparação com bimestre anterior
        let compMediaHtml = '', compFreqHtml = '', mediaAntHtml = '', freqAntHtml = '';
        if (turmaAnterior) {
          const alunoAnt = turmaAnterior.alunos.find((a) => a.nome === aluno.nome);
          if (alunoAnt) {
            const mediaAnt = calcMedia(alunoAnt, turmaAnterior.disciplinas, config);
            const freqAnt = calcFreqMedia(alunoAnt, turmaAnterior.disciplinas);
            compMediaHtml = renderDelta(media, mediaAnt);
            compFreqHtml = renderDelta(freqMedia, freqAnt, true);
            mediaAntHtml = mediaAnt !== null ? mediaAnt.toFixed(1) : '—';
            freqAntHtml = freqAnt !== null ? freqAnt.toFixed(1) + '%' : '—';
          } else {
            mediaAntHtml = '<span class="text-muted">novo</span>';
            freqAntHtml = '<span class="text-muted">—</span>';
          }
        }

        const colsAnt = turmaAnterior ? `
            <td class="table-numeric text-muted">${mediaAntHtml}</td>
            <td class="table-numeric text-muted">${freqAntHtml}</td>` : '';

        return `
          <tr>
            <td>${esc(aluno.nome)}</td>
            <td>${sitBadge}</td>
            <td class="table-numeric font-semibold">${media !== null ? media.toFixed(1) : '—'} ${compMediaHtml}</td>
            <td class="table-numeric">${freqMedia !== null ? freqMedia.toFixed(1) + '%' : '—'} ${compFreqHtml}</td>
            ${colsAnt}
            <td class="table-center">
              <button class="btn btn-ghost btn-sm btn-aluno-detalhe" data-aluno="${esc(aluno.nome)}" aria-label="Ver detalhes de ${esc(aluno.nome)}">
                <i class="bi bi-arrow-right-circle"></i>
              </button>
            </td>
          </tr>`;
      })
      .join('');

    // Bind click events for detail buttons
    tbody.querySelectorAll('.btn-aluno-detalhe').forEach((btn) => {
      btn.addEventListener('click', () => {
        const nome = btn.getAttribute('data-aluno');
        const aluno = alunos.find((a) => a.nome === nome);
        if (aluno) showAlunoModal(aluno, turma, config);
      });
    });
  }

  // -----------------------------------------------------------------------
  // Gráfico 1: Média por Disciplina (horizontal bar)
  // -----------------------------------------------------------------------
  const discLabels = [];
  const discMedias = [];
  const discColors = [];

  for (const disc of turma.disciplinas) {
    let soma = 0;
    let count = 0;
    for (const aluno of alunos) {
      const dados = aluno.disciplinas[disc.nome];
      if (!dados) continue;
      const nota = resolverNota(dados.nota, config);
      if (nota !== null) { soma += nota; count++; }
    }
    if (count > 0) {
      discLabels.push(disc.nome);
      const m = soma / count;
      discMedias.push(m);
      discColors.push(m >= config.params.mediaAprov ? COLORS.success : COLORS.critical);
    }
  }

  criarBarHorizontal('chart-turma-media-disc', {
    labels: discLabels,
    datasets: [
      { label: 'Média', data: discMedias, backgroundColor: discColors, borderRadius: 4, borderSkipped: false },
    ],
    ymax: 10,
    plugins: [pluginBarPct],
  });

  // -----------------------------------------------------------------------
  // Gráfico 2: Frequência Média por Disciplina (bar)
  // -----------------------------------------------------------------------
  const freqLabels = [];
  const freqValues = [];
  const freqColors2 = [];

  for (const disc of turma.disciplinas) {
    if (disc.aulasDadas === 0) continue;
    let somaFaltas = 0;
    let countFreq = 0;
    for (const aluno of alunos) {
      const dados = aluno.disciplinas[disc.nome];
      if (!dados) continue;
      somaFaltas += dados.faltas;
      countFreq++;
    }
    if (countFreq > 0) {
      const freqMediaDisc = ((disc.aulasDadas * countFreq - somaFaltas) / (disc.aulasDadas * countFreq)) * 100;
      freqLabels.push(disc.nome);
      freqValues.push(freqMediaDisc);
      freqColors2.push(freqMediaDisc >= 75 ? COLORS.success : freqMediaDisc >= 50 ? COLORS.warning : COLORS.critical);
    }
  }

  criarBarHorizontal('chart-turma-freq-disc', {
    labels: freqLabels,
    datasets: [
      { label: 'Frequência', data: freqValues, backgroundColor: freqColors2, borderRadius: 4, borderSkipped: false },
    ],
    ymax: 100,
    suffix: '%',
    plugins: [pluginBarPct],
  });
}

// ---------------------------------------------------------------------------
// Modal: Detalhes do Aluno
// ---------------------------------------------------------------------------

function showAlunoModal(aluno, turma, config) {
  const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
  const media = calcMedia(aluno, turma.disciplinas, config);
  const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

  let statusLabel = sit;
  if (sit === 'Aprovado') {
    if (nivelFreq === 'alerta') statusLabel = 'Aprovado ⚠ Alerta de frequência';
    else if (nivelFreq === 'critico') statusLabel = 'Aprovado ⚠ Frequência crítica';
  }
  const statusClass = sit === 'Aprovado'
    ? (nivelFreq ? 'text-warning' : 'text-success')
    : 'text-critical';

  let rows = '';
  for (const disc of turma.disciplinas) {
    const dados = aluno.disciplinas[disc.nome];
    if (!dados) continue;
    const nota = dados.nota !== null
      ? (typeof dados.nota === 'number' ? dados.nota.toFixed(1) : dados.nota)
      : '—';
    const freq = calcFreq(dados.faltas, disc.aulasDadas);
    const freqStr = freq !== null ? freq.toFixed(1) + '%' : '—';
    const acStr = dados.ac > 0 ? `<span class="badge badge-neutral">AC: ${dados.ac}</span>` : '';

    const notaResolvida = resolverNota(dados.nota, config);
    const abaixoMedia = notaResolvida !== null && notaResolvida < config.params.mediaAprov;

    rows += `
      <tr>
        <td>${esc(disc.nome)}</td>
        <td class="table-numeric ${abaixoMedia ? 'text-critical font-semibold' : ''}">${nota} ${acStr}</td>
        <td class="table-numeric">${dados.faltas}</td>
        <td class="table-numeric">${freqStr}</td>
      </tr>`;
  }

  const bodyHtml = `
    <div class="grid-2 mb-4">
      <div class="stat-card">
        <span class="stat-card-label">Situação</span>
        <span class="stat-card-value ${statusClass}">${statusLabel}</span>
      </div>
      <div class="stat-card">
        <span class="stat-card-label">Média / Frequência</span>
        <span class="stat-card-value">${media !== null ? media.toFixed(1) : '—'} / ${freqMedia !== null ? freqMedia.toFixed(1) + '%' : '—'}</span>
      </div>
    </div>
    <div class="table-wrapper">
      <table class="table table-compact">
        <thead>
          <tr>
            <th>Disciplina</th>
            <th class="table-numeric">Nota</th>
            <th class="table-numeric">Faltas</th>
            <th class="table-numeric">Frequência</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  document.getElementById('modal-aluno-title').textContent = aluno.nome;
  showModal(MODAL_ID, bodyHtml);

  // Mostrar botão de imprimir e configurar handler
  const btnPrint = document.getElementById('modal-aluno-btn-print');
  if (btnPrint) {
    btnPrint.style.display = '';
    btnPrint.onclick = () => imprimirBoletim(aluno, turma, config);
  }
}

// ---------------------------------------------------------------------------
// Impressão do boletim
// ---------------------------------------------------------------------------

function imprimirBoletim(aluno, turma, config) {
  const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
  const media = calcMedia(aluno, turma.disciplinas, config);
  const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

  let statusLabel = sit;
  if (sit === 'Aprovado' && nivelFreq === 'alerta') statusLabel = 'Aprovado ⚠ Alerta de frequência';
  else if (sit === 'Aprovado' && nivelFreq === 'critico') statusLabel = 'Aprovado ⚠ Frequência crítica';

  let discRows = '';
  for (const disc of turma.disciplinas) {
    const dados = aluno.disciplinas[disc.nome];
    if (!dados) continue;
    const nota = dados.nota !== null
      ? (typeof dados.nota === 'number' ? dados.nota.toFixed(1) : dados.nota)
      : '—';
    const freq = calcFreq(dados.faltas, disc.aulasDadas);
    const freqStr = freq !== null ? freq.toFixed(1) + '%' : '—';

    discRows += `<tr>
      <td>${disc.nomeOriginal || disc.nome}</td>
      <td style="text-align:center;">${nota}</td>
      <td style="text-align:center;">${dados.faltas}</td>
      <td style="text-align:center;">${freqStr}</td>
    </tr>`;
  }

  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>Boletim — ${aluno.nome}</title>
  <style>
    @page { margin: 15mm; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Georgia', 'Times New Roman', serif;
      font-size: 13px;
      color: #131C2C;
      line-height: 1.5;
    }
    .header {
      text-align: center;
      border-bottom: 2px solid #131C2C;
      padding-bottom: 12px;
      margin-bottom: 16px;
    }
    .header h1 { font-size: 18px; margin-bottom: 4px; }
    .header h2 { font-size: 14px; font-weight: normal; color: #5A6575; }
    .info {
      display: flex;
      justify-content: space-between;
      margin-bottom: 12px;
      padding: 10px;
      background: #F5F3EF;
      border-radius: 4px;
    }
    .info-box { text-align: center; }
    .info-box strong { font-size: 18px; display: block; }
    .info-box small { font-size: 11px; color: #8B95A3; }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 8px;
    }
    th {
      background: #2A3F5A;
      color: white;
      font-size: 10px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      padding: 6px 8px;
      text-align: left;
    }
    td {
      padding: 5px 8px;
      border-bottom: 1px solid #E0DCD5;
      font-size: 12px;
    }
    tr:nth-child(even) td { background: #FAF8F5; }
    .aprovado { color: #4D8C62; font-weight: bold; }
    .reprovado { color: #C94A44; font-weight: bold; }
    .abaixo { color: #C94A44; font-weight: bold; }
    .footer {
      margin-top: 20px;
      font-size: 10px;
      color: #8B95A3;
      text-align: center;
      border-top: 1px solid #E0DCD5;
      padding-top: 10px;
    }
    @media print {
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>${config.escola || 'Conselho de Classe'}</h1>
    <h2>Boletim Escolar — ${turma.bimestre}º Bimestre</h2>
  </div>
  <div class="info">
    <div class="info-box">
      <small>Aluno(a)</small>
      <strong>${aluno.nome}</strong>
    </div>
    <div class="info-box">
      <small>Turma</small>
      <strong>${turma.chave}</strong>
    </div>
    <div class="info-box">
      <small>Média</small>
      <strong>${media !== null ? media.toFixed(1) : '—'}</strong>
    </div>
    <div class="info-box">
      <small>Frequência</small>
      <strong>${freqMedia !== null ? freqMedia.toFixed(1) + '%' : '—'}</strong>
    </div>
    <div class="info-box">
      <small>Situação</small>
      <strong class="${sit === 'Aprovado' ? 'aprovado' : 'reprovado'}">${statusLabel}</strong>
    </div>
  </div>
  <table>
    <thead>
      <tr>
        <th>Disciplina</th>
        <th style="text-align:center; width:60px;">Nota</th>
        <th style="text-align:center; width:60px;">Faltas</th>
        <th style="text-align:center; width:80px;">Frequência</th>
      </tr>
    </thead>
    <tbody>${discRows}</tbody>
  </table>
  <div class="footer">
    Gerado em ${new Date().toLocaleDateString('pt-BR')} — Dashboard Conselho de Classe
  </div>
  <script>window.onload = () => window.print();</script>
</body>
</html>`;

  const w = window.open('', '_blank', 'width=800,height=600');
  if (w) {
    w.document.write(html);
    w.document.close();
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function populateSelector(select) {
  if (!select) return;
  const current = select.value;
  const chaves = getTurmaChaves();
  select.innerHTML = chaves
    .map((c) => `<option value="${c}" ${c === current ? 'selected' : ''}>Turma ${c}</option>`)
    .join('');
  if (chaves.length > 0 && !chaves.includes(current)) {
    select.value = chaves[0];
  }
}

function clearTurma() {
  document.getElementById('turma-alunos').textContent = '—';
  document.getElementById('turma-media').textContent = '—';
  document.getElementById('turma-aprovados').textContent = '—';
  document.getElementById('turma-alerta-freq').textContent = '—';
  const tbody = document.getElementById('turma-tabela-body');
  if (tbody) tbody.innerHTML = '';
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function renderDelta(atual, anterior, pct = false) {
  if (atual === null || anterior === null) return '';
  const diff = atual - anterior;
  if (Math.abs(diff) < 0.1) return '<span style="color:#8B95A3;">→</span>';
  const melhorou = pct ? diff > 0 : diff > 0;
  const arrow = melhorou ? '↑' : '↓';
  const color = melhorou ? '#4D8C62' : '#C94A44';
  const valor = pct ? Math.abs(diff).toFixed(0) + '%' : Math.abs(diff).toFixed(1);
  return `<span style="color:${color}; font-size:0.85em;">${arrow}${valor}</span>`;
}
