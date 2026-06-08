/**
 * config.js — Gerenciamento de configuração via localStorage.
 *
 * Chave persistente: "cc_config" (nunca alterar).
 * Sem dependências de DOM — pode ser importado por calc.js.
 */

import { getTodasDisciplinas, getTurmas, setConfig } from './store.js';

const STORAGE_KEY = 'cc_config';

// ---------------------------------------------------------------------------
// Config padrão
// ---------------------------------------------------------------------------

export function defaultConfig() {
  return {
    areas: {},
    mencoes: {},
    params: {
      mediaAprov: 5.0,
      freqMin: 75,
      freqCrit: 50,
    },
    escola: 'E.E. Professora Célia Vasques Ferrari Duch',
  };
}

// ---------------------------------------------------------------------------
// Áreas disponíveis
// ---------------------------------------------------------------------------

export const AREAS_OPCOES = [
  { value: 'NÃO MAPEADA',                     label: '(Não classificada)' },
  { value: 'Matemática e Ciências da Natureza', label: 'Matemática e Ciências da Natureza' },
  { value: 'Linguagens e Códigos',             label: 'Linguagens e Códigos' },
  { value: 'Ciências Humanas',                 label: 'Ciências Humanas' },
  { value: 'Ensino Técnico / Diversificada',   label: 'Ensino Técnico / Diversificada' },
  { value: 'Projeto de Vida / Eletivas',       label: 'Projeto de Vida / Eletivas' },
];

// ---------------------------------------------------------------------------
// localStorage
// ---------------------------------------------------------------------------

export function loadConfig() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      const cfg = { ...defaultConfig(), ...parsed };
      cfg.params = { ...defaultConfig().params, ...(parsed.params || {}) };
      cfg.mencoes = { ...defaultConfig().mencoes, ...(parsed.mencoes || {}) };
      cfg.areas = { ...defaultConfig().areas, ...(parsed.areas || {}) };
      return cfg;
    }
  } catch {
    // corrupted config — fall through to default
  }
  return defaultConfig();
}

export function saveConfig(config) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
    setConfig(config);
    return true;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// Merge pós-upload
// ---------------------------------------------------------------------------

/**
 * Adiciona disciplinas novas com área "NÃO MAPEADA".
 * Mantém classificações já existentes.
 */
export function mergeDiscs(config) {
  const discNames = getTodasDisciplinas();
  let changed = false;
  for (const nome of discNames) {
    if (!(nome in config.areas)) {
      config.areas[nome] = 'NÃO MAPEADA';
      changed = true;
    }
  }
  return changed;
}

/**
 * Adiciona menções novas com equivalência null (sem valor).
 * Mantém equivalências já configuradas.
 */
export function mergeMencoes(config) {
  const mencoesSet = new Set();
  for (const turma of Object.values(getTurmas())) {
    for (const aluno of turma.alunos) {
      for (const [, dados] of Object.entries(aluno.disciplinas)) {
        if (typeof dados.nota === 'string') {
          mencoesSet.add(dados.nota);
        }
      }
    }
  }
  let changed = false;
  for (const menc of mencoesSet) {
    if (!(menc in config.mencoes)) {
      config.mencoes[menc] = null;
      changed = true;
    }
  }
  return changed;
}

/**
 * Verifica se há disciplinas com área "NÃO MAPEADA".
 */
export function hasNaoMapeadas(config) {
  return Object.values(config.areas).some((area) => area === 'NÃO MAPEADA');
}
