/**
 * app.js — Entry point.
 *
 * Inicializa store, config, ui, e roteamento de seções.
 */

import { loadConfig, hasNaoMapeadas } from './config.js';
import { setConfig } from './store.js';
import {
  initSidebar,
  navigateTo,
  initLgpd,
  initModalCloseButtons,
  setTopbarStatus,
  initSearch,
  initFullscreen,
  clearSearch,
} from './ui.js';
import { initUpload, atualizarStatusTopbar } from './sections/upload.js';
import { renderGeral } from './sections/geral.js';
import { initTurma, renderTurma } from './sections/turma.js';
import { initDisciplina, renderDisciplina } from './sections/disciplina.js';
import { renderArea } from './sections/area.js';
import { renderMedalhistas } from './sections/medalhistas.js';
import { renderDigital } from './sections/digital.js';
import { initRadar, renderRadar } from './sections/radar.js';
import { initConfiguracoes, renderConfiguracoes } from './sections/configuracoes.js';
import { initPDF } from './pdf.js';
import { initPPTX } from './pptx.js';

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
  // 1. Config
  const config = loadConfig();
  setConfig(config);

  // 2. UI
  initSidebar(onSectionChange);
  initLgpd();
  initModalCloseButtons('modal-aluno-overlay');

  // 3. Upload
  initUpload();

  // 4. Section inits (populate selectors, bind events)
  initTurma();
  initDisciplina();
  initRadar();
  initConfiguracoes();

  // 5. Busca global + modo apresentação
  initSearch();
  initFullscreen();

  // 5. Exportações
  initPDF();
  initPPTX();

  // 5. Status inicial
  setTopbarStatus('Nenhum dado carregado — importe os mapões');

  // 6. Ouvir mudanças de dados
  window.addEventListener('cc:data-changed', onDataChanged);
});

// ---------------------------------------------------------------------------
// Section navigation callback
// ---------------------------------------------------------------------------

function onSectionChange(name) {
  clearSearch();
  switch (name) {
    case 'importar':
      atualizarStatusTopbar();
      break;
    case 'geral':
      renderGeral();
      break;
    case 'turma':
      renderTurma();
      break;
    case 'disciplina':
      renderDisciplina();
      break;
    case 'area':
      renderArea();
      break;
    case 'medalhistas':
      renderMedalhistas();
      break;
    case 'digital':
      renderDigital();
      break;
    case 'radar':
      renderRadar();
      break;
    case 'configuracoes':
      renderConfiguracoes();
      break;
    case 'relatorios':
    case 'apresentacoes':
      // Export buttons are always ready via init
      break;
  }
}

// ---------------------------------------------------------------------------
// Data changed callback
// ---------------------------------------------------------------------------

function onDataChanged() {
  // Re-renderizar seção ativa
  const active = document.querySelector('.section.active');
  if (active) {
    const name = active.getAttribute('data-section');
    if (name) onSectionChange(name);
  }
}
