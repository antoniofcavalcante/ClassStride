/**
 * sections/configuracoes.js — Configuração de áreas, menções e parâmetros.
 */

import { hasTurmas, getTurmaChaves, getTurmas, getConfig } from '../store.js';
import { saveConfig, mergeDiscs, mergeMencoes, AREAS_OPCOES } from '../config.js';

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

export function initConfiguracoes() {
  const saveBtn = document.getElementById('btn-config-salvar');
  if (saveBtn) {
    saveBtn.addEventListener('click', handleSave);
  }

  // Atualizar quando novos dados chegarem
  window.addEventListener('cc:data-changed', () => renderConfiguracoes());
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderConfiguracoes() {
  const config = getConfig();
  const saveBar = document.getElementById('config-save-bar');

  if (!hasTurmas()) {
    // Mostrar mensagem de vazio
    renderAreasTable([], config);
    renderMencoesTable({}, config);
    if (saveBar) saveBar.classList.add('hidden');
    return;
  }

  // Garantir merge de novas disciplinas e menções
  let changed = false;
  if (mergeDiscs(config)) changed = true;
  if (mergeMencoes(config)) changed = true;
  if (changed) saveConfig(config);

  // Preencher tabelas
  const discNames = getAllDisciplinas();
  renderAreasTable(discNames, config);

  const mencoes = getAllMencoes();
  renderMencoesTable(mencoes, config);

  // Preencher parâmetros
  renderParams(config);

  // Mostrar barra de salvar
  if (saveBar) saveBar.classList.remove('hidden');

  // Esconder status de save
  const statusEl = document.getElementById('config-save-status');
  if (statusEl) statusEl.innerHTML = '';
}

// ---------------------------------------------------------------------------
// Áreas table
// ---------------------------------------------------------------------------

function renderAreasTable(discNames, config) {
  const tbody = document.getElementById('config-areas-body');
  const countEl = document.getElementById('config-areas-count');
  if (!tbody) return;

  if (countEl) countEl.textContent = discNames.length + ' disciplina(s)';

  if (discNames.length === 0) {
    tbody.innerHTML = `<tr><td colspan="2" class="text-center text-muted">Nenhuma disciplina carregada. Importe os mapões primeiro.</td></tr>`;
    return;
  }

  const optionsHtml = AREAS_OPCOES.map(
    (opt) => `<option value="${esc(opt.value)}">${esc(opt.label)}</option>`
  ).join('');

  tbody.innerHTML = discNames
    .sort()
    .map((nome) => {
      const current = config.areas[nome] || 'NÃO MAPEADA';
      const rowClass = current === 'NÃO MAPEADA' ? 'style="background: var(--color-amber-light);"' : '';
      return `
        <tr ${rowClass}>
          <td class="font-medium">${esc(nome)}</td>
          <td>
            <select class="form-select config-area-select" data-disciplina="${esc(nome)}" aria-label="Área de ${esc(nome)}">
              ${optionsHtml.replace(`value="${esc(current)}"`, `value="${esc(current)}" selected`)}
            </select>
          </td>
        </tr>`;
    })
    .join('');
}

// ---------------------------------------------------------------------------
// Menções table
// ---------------------------------------------------------------------------

function renderMencoesTable(mencoes, config) {
  const tbody = document.getElementById('config-mencoes-body');
  if (!tbody) return;

  const keys = Object.keys(mencoes).sort();

  if (keys.length === 0) {
    tbody.innerHTML = `<tr><td colspan="2" class="text-center text-muted">Nenhuma menção detectada. Importe os mapões primeiro.</td></tr>`;
    return;
  }

  tbody.innerHTML = keys
    .map((menc) => {
      const val = config.mencoes[menc];
      const valStr = val !== null && val !== undefined ? val : '';
      return `
        <tr>
          <td class="font-medium font-semibold">${esc(menc)}</td>
          <td>
            <input type="number" class="form-input config-mencao-input"
                   data-mencao="${esc(menc)}" value="${valStr}"
                   min="0" max="10" step="0.5"
                   placeholder="Sem equivalência"
                   aria-label="Equivalência numérica para ${esc(menc)}"
                   style="max-width: 160px;">
          </td>
        </tr>`;
    })
    .join('');
}

// ---------------------------------------------------------------------------
// Parâmetros
// ---------------------------------------------------------------------------

function renderParams(config) {
  const media = document.getElementById('param-media-aprov');
  const freqMin = document.getElementById('param-freq-min');
  const freqCrit = document.getElementById('param-freq-crit');

  if (media) media.value = config.params.mediaAprov;
  if (freqMin) freqMin.value = config.params.freqMin;
  if (freqCrit) freqCrit.value = config.params.freqCrit;
}

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------

function handleSave() {
  const config = getConfig();

  // Coletar áreas
  const areaSelects = document.querySelectorAll('.config-area-select');
  for (const sel of areaSelects) {
    config.areas[sel.getAttribute('data-disciplina')] = sel.value;
  }

  // Coletar menções
  const mencInputs = document.querySelectorAll('.config-mencao-input');
  for (const inp of mencInputs) {
    const menc = inp.getAttribute('data-mencao');
    const raw = inp.value.trim();
    if (raw === '') {
      config.mencoes[menc] = null;
    } else {
      const parsed = parseFloat(raw.replace(',', '.'));
      config.mencoes[menc] = isNaN(parsed) ? null : parsed;
    }
  }

  // Coletar parâmetros
  const mediaEl = document.getElementById('param-media-aprov');
  const freqMinEl = document.getElementById('param-freq-min');
  const freqCritEl = document.getElementById('param-freq-crit');

  if (mediaEl) {
    const v = parseFloat(mediaEl.value);
    if (!isNaN(v) && v >= 0 && v <= 10) config.params.mediaAprov = v;
  }
  if (freqMinEl) {
    const v = parseFloat(freqMinEl.value);
    if (!isNaN(v) && v >= 0 && v <= 100) config.params.freqMin = v;
  }
  if (freqCritEl) {
    const v = parseFloat(freqCritEl.value);
    if (!isNaN(v) && v >= 0 && v <= 100) config.params.freqCrit = v;
  }

  // Salvar
  const saved = saveConfig(config);

  // Feedback
  const statusEl = document.getElementById('config-save-status');
  if (statusEl) {
    if (saved) {
      statusEl.innerHTML = '<i class="bi bi-check-circle-fill"></i> Configurações salvas com sucesso!';
      statusEl.style.color = 'var(--color-sage)';
    } else {
      statusEl.innerHTML = '<i class="bi bi-x-circle-fill"></i> Erro ao salvar configurações.';
      statusEl.style.color = 'var(--color-crimson)';
    }
    // Limpar após 3s
    setTimeout(() => { statusEl.innerHTML = ''; }, 3000);
  }

  // Disparar re-renderização
  window.dispatchEvent(new CustomEvent('cc:data-changed'));
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getAllDisciplinas() {
  const set = new Set();
  for (const turma of Object.values(getTurmas())) {
    for (const disc of turma.disciplinas) {
      set.add(disc.nome);
    }
  }
  return [...set];
}

function getAllMencoes() {
  const set = new Set();
  for (const turma of Object.values(getTurmas())) {
    for (const aluno of turma.alunos) {
      for (const [, dados] of Object.entries(aluno.disciplinas)) {
        if (typeof dados.nota === 'string') {
          set.add(dados.nota);
        }
      }
    }
  }
  return [...set];
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
