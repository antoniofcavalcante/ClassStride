/**
 * sections/disciplina.js — Drill-down por disciplina.
 */

import { getTurma, getTurmaChaves, hasTurmas, getConfig } from '../store.js';
import { calcFreq, resolverNota } from '../calc.js';
import { criarBarChart, pluginBarPct, COLORS } from '../charts.js';

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

export function initDisciplina() {
  const selectTurma = document.getElementById('select-disciplina-turma');
  const selectDisc = document.getElementById('select-disciplina');

  if (selectTurma) {
    selectTurma.addEventListener('change', () => {
      populateDiscSelector(selectTurma.value);
      renderDisciplina();
    });
  }

  if (selectDisc) {
    selectDisc.addEventListener('change', () => renderDisciplina());
  }

  window.addEventListener('cc:data-changed', () => {
    if (selectTurma) populateTurmaSelector(selectTurma);
    renderDisciplina();
  });
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderDisciplina() {
  if (!hasTurmas()) {
    clearDisciplina();
    return;
  }

  const selectTurma = document.getElementById('select-disciplina-turma');
  const selectDisc = document.getElementById('select-disciplina');

  populateTurmaSelector(selectTurma);
  const chave = selectTurma?.value || getTurmaChaves()[0];
  if (selectTurma && selectTurma.value !== chave) selectTurma.value = chave;

  const turma = getTurma(chave);
  if (!turma) { clearDisciplina(); return; }

  populateDiscSelector(chave, selectDisc?.value);

  const discNome = selectDisc?.value;
  if (!discNome) { clearDisciplina(); return; }

  const config = getConfig();
  const discInfo = turma.disciplinas.find((d) => d.nome === discNome);
  if (!discInfo) { clearDisciplina(); return; }

  // -----------------------------------------------------------------------
  // Notas e frequências dos alunos nessa disciplina
  // -----------------------------------------------------------------------
  const notas = [];
  const frequencias = [];
  let somaNota = 0;
  let countNota = 0;
  let abaixoMedia = 0;
  let alertaFreq = 0;

  for (const aluno of turma.alunos) {
    const dados = aluno.disciplinas[discNome];
    if (!dados) continue;

    const nota = resolverNota(dados.nota, config);
    if (nota !== null) {
      notas.push({ aluno: aluno.nome, nota, faltas: dados.faltas, ac: dados.ac });
      somaNota += nota;
      countNota++;
      if (nota < config.params.mediaAprov) abaixoMedia++;
    }

    const freq = calcFreq(dados.faltas, discInfo.aulasDadas);
    if (freq !== null) {
      frequencias.push(freq);
      if (freq < config.params.freqMin) alertaFreq++;
    }
  }

  const mediaGeral = countNota > 0 ? somaNota / countNota : 0;

  // -----------------------------------------------------------------------
  // Cards
  // -----------------------------------------------------------------------
  document.getElementById('disc-media').textContent = countNota > 0 ? mediaGeral.toFixed(1) : '—';
  document.getElementById('disc-abaixo-media').textContent = abaixoMedia;
  document.getElementById('disc-alerta-freq').textContent = alertaFreq;
  document.getElementById('disc-aulas-dadas').textContent = discInfo.aulasDadas;

  // -----------------------------------------------------------------------
  // Distribuição de notas (histograma)
  // -----------------------------------------------------------------------
  if (notas.length > 0) {
    const bins = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    const histLabels = [];
    const histData = [];
    const histColors = [];

    for (let i = 0; i < bins.length - 1; i++) {
      const lo = bins[i];
      const hi = bins[i + 1];
      const label = i < bins.length - 2 ? `${lo}–${hi}` : `${hi}`;
      histLabels.push(label);
      const count = notas.filter((n) => n.nota >= lo && n.nota < hi).length;
      // Include exact 10 in last bin
      const countExact = i === bins.length - 2
        ? count + notas.filter((n) => n.nota === hi).length
        : count;
      histData.push(countExact);
      histColors.push(lo < config.params.mediaAprov ? COLORS.critical : COLORS.success);
    }

    criarBarChart('chart-disciplina-dist', {
      labels: histLabels,
      datasets: [
        {
          label: 'Alunos',
          data: histData,
          backgroundColor: histColors,
          borderRadius: 4,
          borderSkipped: false,
        },
      ],
      ymax: Math.max(...histData, 1) + 1,
      plugins: [{ id: 'countLabels', ...pluginBarPct, afterDatasetsDraw(chart) {
        const { ctx } = chart;
        chart.data.datasets.forEach((ds, i) => {
          const meta = chart.getDatasetMeta(i);
          if (meta.hidden) return;
          meta.data.forEach((bar, idx) => {
            const v = ds.data[idx];
            if (v > 0) {
              ctx.fillStyle = '#131C2C';
              ctx.font = "600 11px 'Source Sans 3', sans-serif";
              ctx.textAlign = 'center';
              ctx.textBaseline = 'middle';
              ctx.fillText(v, bar.x, bar.y + bar.height / 2);
            }
          });
        });
      }}],
    });
  }

  // -----------------------------------------------------------------------
  // Tabela de alunos
  // -----------------------------------------------------------------------
  const tbody = document.getElementById('disciplina-tabela-body');
  if (tbody) {
    tbody.innerHTML = turma.alunos
      .map((aluno) => {
        const dados = aluno.disciplinas[discNome];
        if (!dados) return '';

        const nota = dados.nota !== null
          ? (typeof dados.nota === 'number' ? dados.nota.toFixed(1) : dados.nota)
          : '—';
        const freq = calcFreq(dados.faltas, discInfo.aulasDadas);
        const freqStr = freq !== null ? freq.toFixed(1) + '%' : '—';

        const notaResolvida = resolverNota(dados.nota, config);
        const abaixo = notaResolvida !== null && notaResolvida < config.params.mediaAprov;

        return `
          <tr>
            <td>${esc(aluno.nome)}</td>
            <td class="table-numeric ${abaixo ? 'text-critical font-semibold' : ''}">${nota}</td>
            <td class="table-numeric">${dados.faltas}</td>
            <td class="table-numeric">${freqStr}</td>
          </tr>`;
      })
      .join('');
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function populateTurmaSelector(select) {
  if (!select) return;
  const current = select.value;
  const chaves = getTurmaChaves();
  select.innerHTML = chaves
    .map((c) => `<option value="${c}" ${c === current ? 'selected' : ''}>Turma ${c}</option>`)
    .join('');
  if (chaves.length > 0 && !chaves.includes(current)) select.value = chaves[0];
}

function populateDiscSelector(chave, keepValue) {
  const select = document.getElementById('select-disciplina');
  if (!select) return;

  const turma = getTurma(chave);
  if (!turma) { select.innerHTML = ''; return; }

  const current = keepValue || select.value;
  select.innerHTML = turma.disciplinas
    .map((d) => `<option value="${d.nome}" ${d.nome === current ? 'selected' : ''}>${d.nomeOriginal || d.nome}</option>`)
    .join('');
  if (!turma.disciplinas.find((d) => d.nome === current)) {
    select.value = turma.disciplinas[0]?.nome || '';
  }
}

function clearDisciplina() {
  document.getElementById('disc-media').textContent = '—';
  document.getElementById('disc-abaixo-media').textContent = '—';
  document.getElementById('disc-alerta-freq').textContent = '—';
  document.getElementById('disc-aulas-dadas').textContent = '—';
  const tbody = document.getElementById('disciplina-tabela-body');
  if (tbody) tbody.innerHTML = '';
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
