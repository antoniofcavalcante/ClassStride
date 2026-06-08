/**
 * sections/upload.js — Importação de dados (Mapão + Digital).
 *
 * Gerencia dropzones, parse de arquivos, feedback visual,
 * alerta de disciplinas não-mapeadas.
 */

import {
  hasTurmas,
  getTotalAlunosAtivos,
  getTurmaChaves,
  getTurma,
  getTodasDisciplinas,
  setTurma,
  setTurmaAnterior,
  setDigital,
  getDigital,
  getConfig,
  removeTurma,
  hasTurmasAnterior,
  getTurmaAnterior,
} from '../store.js';
import { mergeDiscs, mergeMencoes, hasNaoMapeadas, saveConfig } from '../config.js';
import { parseMapao, parseDigital } from '../parser.js';
import { setTopbarStatus, showLoader, hideLoader } from '../ui.js';

// ---------------------------------------------------------------------------
// Init: configura event listeners dos dropzones
// ---------------------------------------------------------------------------

export function initUpload() {
  setupDropzone('dropzone-mapa', 'input-mapa', handleMapaoFiles);
  setupDropzone('dropzone-digital', 'input-digital', handleDigitalFiles);
  setupDropzone('dropzone-anterior', 'input-anterior', handleAnteriorFiles);

  // Export Excel button
  const btnExport = document.getElementById('btn-exportar-excel');
  if (btnExport) {
    btnExport.addEventListener('click', exportarExcel);
  }
}

// ---------------------------------------------------------------------------
// Dropzone setup (drag/drop + click + keyboard)
// ---------------------------------------------------------------------------

function setupDropzone(dropzoneId, inputId, handler) {
  const dropzone = document.getElementById(dropzoneId);
  const input = document.getElementById(inputId);
  if (!dropzone || !input) return;

  // Click to open file dialog
  dropzone.addEventListener('click', () => input.click());
  dropzone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      input.click();
    }
  });

  // File input change
  input.addEventListener('change', () => {
    if (input.files.length > 0) {
      const files = Array.from(input.files);
      input.value = '';
      handler(files);
    }
  });

  // Drag & drop
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-over');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      handler(Array.from(e.dataTransfer.files));
    }
  });
}

// ---------------------------------------------------------------------------
// Mapão handler
// ---------------------------------------------------------------------------

async function handleMapaoFiles(files) {
  showLoader('Processando mapões...');

  const resultados = [];
  const erros = [];

  for (const file of files) {
    // Validar extensão
    if (!file.name.match(/\.xlsx?$/i)) {
      erros.push(`${file.name}: formato não suportado`);
      continue;
    }

    try {
      const turmaData = await parseMapao(file);
      resultados.push(turmaData);
    } catch (err) {
      erros.push(`${file.name}: ${err.message}`);
    }
  }

  // Merge no store
  for (const turmaData of resultados) {
    setTurma(turmaData.chave, turmaData);
  }

  // Atualizar config com novas disciplinas/menções
  const config = getConfig();
  let configChanged = false;
  if (mergeDiscs(config)) configChanged = true;
  if (mergeMencoes(config)) configChanged = true;
  if (configChanged) saveConfig(config);

  hideLoader();

  // Feedback
  if (erros.length > 0) {
    const errosDiv = document.getElementById('upload-errors');
    if (errosDiv) {
      errosDiv.innerHTML = erros.map((e) => `<p>${e}</p>`).join('');
      errosDiv.classList.remove('hidden');
    }
  }

  renderUploadSummary();

  // Atualizar status
  atualizarStatusTopbar();

  // Atualizar sidebar
  atualizarSidebarTurmas();

  // Disparar evento para outros módulos
  window.dispatchEvent(new CustomEvent('cc:data-changed'));
}

// ---------------------------------------------------------------------------
// Digital handler
// ---------------------------------------------------------------------------

async function handleDigitalFiles(files) {
  showLoader('Processando planilha digital...');
  const erros = [];

  for (const file of files) {
    if (!file.name.match(/\.xlsx?$/i)) {
      erros.push(`${file.name}: formato não suportado`);
      continue;
    }

    try {
      const rows = await parseDigital(file);
      setDigital([...getDigital(), ...rows]);
    } catch (err) {
      erros.push(`${file.name}: ${err.message}`);
    }
  }

  hideLoader();

  if (erros.length > 0) {
    console.warn('Erros ao processar digital:', erros);
  }

  atualizarStatusTopbar();
  window.dispatchEvent(new CustomEvent('cc:data-changed'));
}

// ---------------------------------------------------------------------------
// Render Summary
// ---------------------------------------------------------------------------

function renderUploadSummary() {
  const summaryEl = document.getElementById('upload-summary');
  if (!summaryEl) return;

  if (!hasTurmas()) {
    summaryEl.classList.add('hidden');
    return;
  }

  summaryEl.classList.remove('hidden');

  // Totals
  const totalAlunos = getTotalAlunosAtivos();
  const totalDiscs = getTodasDisciplinas().length;
  const totalTurmas = getTurmaChaves().length;

  const elTurmas = document.getElementById('summary-turmas');
  const elAlunos = document.getElementById('summary-alunos');
  const elDiscs = document.getElementById('summary-disciplinas');

  if (elTurmas) elTurmas.textContent = totalTurmas;
  if (elAlunos) elAlunos.textContent = totalAlunos;
  if (elDiscs) elDiscs.textContent = totalDiscs;

  // Alert for unmapped
  const alertEl = document.getElementById('alert-nao-mapeadas');
  if (alertEl) {
    const config = getConfig();
    if (hasNaoMapeadas(config)) {
      alertEl.classList.remove('hidden');
    } else {
      alertEl.classList.add('hidden');
    }
  }

  // Turma chips com botão de excluir
  const chipsEl = document.getElementById('turmas-chips');
  if (chipsEl) {
    const chaves = getTurmaChaves();
    chipsEl.innerHTML = chaves
      .map((chave) => {
        const turma = getTurma(chave);
        const count = turma ? turma.alunos.length : 0;
        return `<span class="upload-turma-chip">
          ${chave} · ${count} alunos
          <button class="chip-remove" data-remove-turma="${chave}" title="Remover ${chave}" aria-label="Remover turma ${chave}">✕</button>
        </span>`;
      })
      .join('');

    // Bind delete handlers
    chipsEl.querySelectorAll('[data-remove-turma]').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const chave = btn.getAttribute('data-remove-turma');
        if (chave && confirm(`Remover turma ${chave} e todos os seus dados?`)) {
          removeTurma(chave);
          renderUploadSummary();
          atualizarStatusTopbar();
          window.dispatchEvent(new CustomEvent('cc:data-changed'));
        }
      });
    });
  }
}

// ---------------------------------------------------------------------------
// Sidebar update (add turma badges)
// ---------------------------------------------------------------------------

function atualizarSidebarTurmas() {
  // Atualizar nome da escola no footer
  const config = getConfig();
  const schoolEl = document.getElementById('sidebar-school-name');
  if (schoolEl && config.escola) {
    schoolEl.textContent = config.escola;
  }
}

// ---------------------------------------------------------------------------
// Topbar status
// ---------------------------------------------------------------------------

export function atualizarStatusTopbar() {
  if (hasTurmas()) {
    const totalTurmas = getTurmaChaves().length;
    const totalAlunos = getTotalAlunosAtivos();
    const extra = hasTurmasAnterior() ? ' · com bimestre anterior' : '';
    setTopbarStatus(`${totalTurmas} turma(s) · ${totalAlunos} alunos ativos${extra}`);
  } else {
    setTopbarStatus('Nenhum dado carregado');
  }
}

// ---------------------------------------------------------------------------
// Bimestre anterior handler
// ---------------------------------------------------------------------------

async function handleAnteriorFiles(files) {
  const filesArr = Array.from(files);
  showLoader('Processando mapões do bimestre anterior...');
  const erros = [];

  for (const file of filesArr) {
    if (!file.name.match(/\.xlsx?$/i)) {
      erros.push(`${file.name}: formato não suportado`);
      continue;
    }
    try {
      const turmaData = await parseMapao(file);
      setTurmaAnterior(turmaData.chave, turmaData);
    } catch (err) {
      erros.push(`${file.name}: ${err.message}`);
    }
  }

  hideLoader();
  if (erros.length > 0) {
    console.warn('Erros ao processar bimestre anterior:', erros);
  }

  atualizarStatusTopbar();
  window.dispatchEvent(new CustomEvent('cc:data-changed'));
}

// ---------------------------------------------------------------------------
// Exportar dados processados para Excel
// ---------------------------------------------------------------------------

function exportarExcel() {
  if (!hasTurmas()) {
    alert('Importe os mapões antes de exportar.');
    return;
  }

  const config = getConfig();
  const chaves = getTurmaChaves();
  const rows = [['Turma', 'Aluno', 'Média', 'Frequência', 'Situação']];

  // Add discipline columns
  const todasDiscs = getTodasDisciplinas();
  const headerDiscs = [];
  for (const disc of todasDiscs) {
    headerDiscs.push(disc + ' Nota', disc + ' Faltas', disc + ' Freq.');
  }
  rows[0].push(...headerDiscs);

  for (const chave of chaves) {
    const turma = getTurma(chave);
    for (const aluno of turma.alunos) {
      const media = calcMediaLocal(aluno, turma, config);
      const freqMedia = calcFreqMediaLocal(aluno, turma);
      const { sit, nivelFreq } = calcSituacaoLocal(aluno, turma, config);

      let sitLabel = sit;
      if (sit === 'Aprovado' && nivelFreq === 'alerta') sitLabel = 'Aprovado (alerta freq)';
      else if (sit === 'Aprovado' && nivelFreq === 'critico') sitLabel = 'Aprovado (freq crítica)';

      const row = [
        chave,
        aluno.nome,
        media !== null ? media.toFixed(1) : '',
        freqMedia !== null ? freqMedia.toFixed(1) + '%' : '',
        sitLabel,
      ];

      for (const discNome of todasDiscs) {
        const discInfo = turma.disciplinas.find((d) => d.nome === discNome);
        const dados = aluno.disciplinas[discNome];
        if (dados && discInfo) {
          const nota = dados.nota !== null ? (typeof dados.nota === 'number' ? dados.nota.toFixed(1) : dados.nota) : '';
          const freq = discInfo.aulasDadas > 0 ? (((discInfo.aulasDadas - dados.faltas) / discInfo.aulasDadas) * 100).toFixed(1) + '%' : '';
          row.push(nota, String(dados.faltas), freq);
        } else {
          row.push('', '', '');
        }
      }
      rows.push(row);
    }
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Dados Processados');
  XLSX.writeFile(wb, 'conselho-dados-processados.xlsx');
}

// Mini versões das funções de calc para evitar import circular
function calcMediaLocal(aluno, turma, config) {
  let soma = 0, count = 0;
  for (const disc of turma.disciplinas) {
    const dados = aluno.disciplinas[disc.nome];
    if (!dados) continue;
    const nota = typeof dados.nota === 'number' ? dados.nota : null;
    if (nota !== null) { soma += nota; count++; }
  }
  return count > 0 ? soma / count : null;
}

function calcFreqMediaLocal(aluno, turma) {
  let totalAulas = 0, totalPresente = 0;
  for (const disc of turma.disciplinas) {
    const dados = aluno.disciplinas[disc.nome];
    if (!dados || disc.aulasDadas === 0) continue;
    totalAulas += disc.aulasDadas;
    totalPresente += (disc.aulasDadas - dados.faltas);
  }
  return totalAulas > 0 ? (totalPresente / totalAulas) * 100 : null;
}

function calcSituacaoLocal(aluno, turma, config) {
  let reprovado = false, temNota = false, minFreq = null;
  for (const disc of turma.disciplinas) {
    const dados = aluno.disciplinas[disc.nome];
    if (!dados) continue;
    const nota = typeof dados.nota === 'number' ? dados.nota : null;
    if (nota !== null) {
      temNota = true;
      if (nota < config.params.mediaAprov) reprovado = true;
    }
    if (disc.aulasDadas > 0) {
      const freq = ((disc.aulasDadas - dados.faltas) / disc.aulasDadas) * 100;
      if (minFreq === null || freq < minFreq) minFreq = freq;
    }
  }
  const sit = reprovado ? 'Reprovado' : 'Aprovado';
  let nivelFreq = null;
  if (minFreq !== null) {
    if (minFreq < config.params.freqCrit) nivelFreq = 'critico';
    else if (minFreq < config.params.freqMin) nivelFreq = 'alerta';
  }
  return { sit, nivelFreq };
}
