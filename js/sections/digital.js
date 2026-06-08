/**
 * sections/digital.js — Progresso das aulas digitais por turma e disciplina.
 */

import { hasTurmas, hasDigital, getTurmaChaves, getDigital } from '../store.js';

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderDigital() {
  const contentEl = document.getElementById('digital-content');
  const emptyEl = document.getElementById('digital-empty');

  if (!contentEl || !emptyEl) return;

  // Filtro de turma
  const selectTurma = document.getElementById('select-digital-turma');
  if (selectTurma && selectTurma.options.length <= 1 && hasTurmas()) {
    const chaves = getTurmaChaves();
    selectTurma.innerHTML = '<option value="todas">Todas as turmas</option>' +
      chaves.map((c) => `<option value="${c}">Turma ${c}</option>`).join('');
    selectTurma.addEventListener('change', () => renderDigital());
  }

  // Filtro de bimestre
  const selectBim = document.getElementById('select-digital-bim');
  if (selectBim && !selectBim._bound) {
    selectBim.addEventListener('change', () => renderDigital());
    selectBim._bound = true;
  }

  const filtroTurma = selectTurma?.value || 'todas';
  const bimIdx = parseInt(selectBim?.value || '1', 10) - 1;

  if (!hasDigital()) {
    contentEl.innerHTML = '';
    emptyEl.classList.remove('hidden');
    return;
  }

  emptyEl.classList.add('hidden');
  const digital = getDigital();
  const chavesTurmas = hasTurmas() ? getTurmaChaves() : [];
  const turmasMostrar = filtroTurma === 'todas' ? chavesTurmas : [filtroTurma];

  // Agrupar digital por turma
  const porTurma = {};
  for (const row of digital) {
    if (!porTurma[row.chave]) porTurma[row.chave] = [];
    porTurma[row.chave].push(row);
  }

  let html = '';

  for (const chave of turmasMostrar) {
    const rows = porTurma[chave];
    if (!rows || rows.length === 0) continue;

    html += `<div class="digital-turma-group">
      <h3 class="digital-turma-title">Turma ${chave} — ${bimIdx + 1}º Bimestre</h3>`;

    for (const row of rows) {
      const bimData = row.bimestres[bimIdx];
      const previsto = bimData?.previsto || 0;
      const concluido = bimData?.concluido || 0;
      const pct = previsto > 0 ? Math.round((concluido / previsto) * 100) : 0;
      const fillClass = pct >= 100 ? '' : pct >= 50 ? 'warning' : 'critical';

      html += `<div class="card mb-2">
        <div class="card-body" style="padding: var(--space-4) var(--space-5);">
          <div class="progress-label">
            <span class="font-semibold">${esc(row.disciplina)}</span>
            <span>${concluido} / ${previsto} aulas</span>
          </div>
          <div class="progress">
            <div class="progress-fill ${fillClass}" style="width: ${pct}%;"></div>
          </div>
        </div>
      </div>`;
    }

    html += `</div>`;
  }

  // Turmas com dados digitais mas sem mapão carregado
  for (const [chave, rows] of Object.entries(porTurma)) {
    if (chavesTurmas.includes(chave)) continue;
    if (hasTurmas()) {
      html += `<div class="alert alert-info mt-4" role="alert">
        <i class="bi bi-info-circle-fill"></i>
        <span>Planilha digital contém dados da turma <strong>${chave}</strong>, mas o mapão dessa turma não foi carregado.</span>
      </div>`;
    }
  }

  contentEl.innerHTML = html || '<p class="text-muted text-center">Nenhum dado para exibir.</p>';
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
