/**
 * sections/medalhistas.js — Destaques positivos
 * (melhor média por turma, Ouro/Prata/Bronze, frequência perfeita).
 */

import { getTurmaChaves, getTurma, hasTurmas, getConfig } from '../store.js';
import { calcMedia, calcFreqMedia } from '../calc.js';

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderMedalhistas() {
  if (!hasTurmas()) {
    showAllEmpty();
    return;
  }

  const config = getConfig();
  const chaves = getTurmaChaves();

  // Filtro de turma
  const select = document.getElementById('select-medalhistas-turma');
  if (select && select.options.length <= 1) {
    select.innerHTML = '<option value="todas">Todas as turmas</option>' +
      chaves.map((c) => `<option value="${c}">Turma ${c}</option>`).join('');
    select.addEventListener('change', () => renderMedalhistas());
  }
  const filtroTurma = select?.value || 'todas';
  const turmasFiltradas = filtroTurma === 'todas' ? chaves : [filtroTurma];

  // Coletar todos os alunos com métricas
  const todosAlunos = [];
  for (const chave of turmasFiltradas) {
    const turma = getTurma(chave);
    for (const aluno of turma.alunos) {
      const media = calcMedia(aluno, turma.disciplinas, config);
      const freq = calcFreqMedia(aluno, turma.disciplinas);
      if (media !== null && freq !== null) {
        todosAlunos.push({ nome: aluno.nome, turma: chave, media, freq });
      }
    }
  }

  // -----------------------------------------------------------------------
  // 1. Melhor média por turma
  // -----------------------------------------------------------------------
  const topPorTurma = [];
  for (const chave of turmasFiltradas) {
    const daTurma = todosAlunos.filter((a) => a.turma === chave);
    if (daTurma.length > 0) {
      daTurma.sort((a, b) => b.media - a.media);
      topPorTurma.push(daTurma[0]);
    }
  }
  topPorTurma.sort((a, b) => b.media - a.media);

  renderCardGrid('medalhistas-media-grid', 'medalhistas-media-empty', topPorTurma, (item, i) => `
    <div class="medalhista-card">
      <div class="medalhista-rank">${i + 1}º</div>
      <div class="medalhista-info">
        <div class="nome">${esc(item.nome)}</div>
        <div class="turma">Turma ${item.turma}</div>
      </div>
      <div class="medalhista-info" style="text-align:right;">
        <div class="valor">${item.media.toFixed(1)}</div>
      </div>
    </div>
  `);

  // -----------------------------------------------------------------------
  // 2. Ouro: Freq ≥ 90% · Média ≥ 9
  // -----------------------------------------------------------------------
  const ouro = todosAlunos
    .filter((a) => a.freq >= 90 && a.media >= 9)
    .sort((a, b) => b.media - a.media);

  renderCardGrid('medalhistas-ouro-grid', 'medalhistas-ouro-empty', ouro, (item, i) => `
    <div class="medalhista-card" style="border-left: 4px solid #C4953A;">
      <div class="medalhista-rank" style="color: #C4953A;">🥇</div>
      <div class="medalhista-info">
        <div class="nome">${esc(item.nome)}</div>
        <div class="turma">Turma ${item.turma}</div>
      </div>
      <div class="medalhista-info" style="text-align:right;">
        <div class="valor">${item.media.toFixed(1)}</div>
        <div class="turma">${item.freq.toFixed(0)}% freq.</div>
      </div>
    </div>
  `);

  // -----------------------------------------------------------------------
  // 3. Prata: Freq ≥ 90% · Média 8–9
  // -----------------------------------------------------------------------
  const prata = todosAlunos
    .filter((a) => a.freq >= 90 && a.media >= 8 && a.media < 9)
    .sort((a, b) => b.media - a.media);

  renderCardGrid('medalhistas-prata-grid', 'medalhistas-prata-empty', prata, (item, i) => `
    <div class="medalhista-card" style="border-left: 4px solid #7A8599;">
      <div class="medalhista-rank" style="color: #7A8599;">🥈</div>
      <div class="medalhista-info">
        <div class="nome">${esc(item.nome)}</div>
        <div class="turma">Turma ${item.turma}</div>
      </div>
      <div class="medalhista-info" style="text-align:right;">
        <div class="valor">${item.media.toFixed(1)}</div>
        <div class="turma">${item.freq.toFixed(0)}% freq.</div>
      </div>
    </div>
  `);

  // -----------------------------------------------------------------------
  // 4. Bronze: Freq ≥ 90% · Média 7–8
  // -----------------------------------------------------------------------
  const bronze = todosAlunos
    .filter((a) => a.freq >= 90 && a.media >= 7 && a.media < 8)
    .sort((a, b) => b.media - a.media);

  renderCardGrid('medalhistas-bronze-grid', 'medalhistas-bronze-empty', bronze, (item, i) => `
    <div class="medalhista-card" style="border-left: 4px solid #C26547;">
      <div class="medalhista-rank" style="color: #C26547;">🥉</div>
      <div class="medalhista-info">
        <div class="nome">${esc(item.nome)}</div>
        <div class="turma">Turma ${item.turma}</div>
      </div>
      <div class="medalhista-info" style="text-align:right;">
        <div class="valor">${item.media.toFixed(1)}</div>
        <div class="turma">${item.freq.toFixed(0)}% freq.</div>
      </div>
    </div>
  `);

  // -----------------------------------------------------------------------
  // 5. Frequência perfeita (100%)
  // -----------------------------------------------------------------------
  const freqPerfeita = todosAlunos
    .filter((a) => a.freq >= 100)
    .sort((a, b) => a.nome.localeCompare(b.nome));

  renderCardGrid('medalhistas-freq-grid', 'medalhistas-freq-empty', freqPerfeita, (item) => `
    <div class="medalhista-card">
      <div class="medalhista-rank" style="color: #4D8C62;"><i class="bi bi-star-fill"></i></div>
      <div class="medalhista-info">
        <div class="nome">${esc(item.nome)}</div>
        <div class="turma">Turma ${item.turma}</div>
      </div>
      <div class="medalhista-info" style="text-align:right;">
        <div class="valor">100%</div>
      </div>
    </div>
  `);
}

// ---------------------------------------------------------------------------
// Helper: renderiza um grid ou empty state
// ---------------------------------------------------------------------------

function renderCardGrid(gridId, emptyId, items, templateFn) {
  const grid = document.getElementById(gridId);
  const empty = document.getElementById(emptyId);

  if (!grid) return;

  if (items.length > 0) {
    grid.innerHTML = items.map(templateFn).join('');
    grid.classList.remove('hidden');
    if (empty) empty.classList.add('hidden');
  } else {
    grid.classList.add('hidden');
    if (empty) empty.classList.remove('hidden');
  }
}

// ---------------------------------------------------------------------------
// Show all empty
// ---------------------------------------------------------------------------

function showAllEmpty() {
  const ids = [
    'medalhistas-media-grid', 'medalhistas-media-empty',
    'medalhistas-ouro-grid', 'medalhistas-ouro-empty',
    'medalhistas-prata-grid', 'medalhistas-prata-empty',
    'medalhistas-bronze-grid', 'medalhistas-bronze-empty',
    'medalhistas-freq-grid', 'medalhistas-freq-empty',
  ];
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) {
      if (id.includes('empty')) el.classList.remove('hidden');
      else el.classList.add('hidden');
    }
  }
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
