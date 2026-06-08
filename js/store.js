/**
 * store.js — Estado global único da aplicação.
 *
 * ÚNICO módulo que manipula o objeto APP diretamente.
 * Todos os outros módulos importam os getters/actions daqui.
 */

export const APP = {
  turmas: {},
  turmasAnterior: {},
  digital: [],
  config: null,
};

// ---------------------------------------------------------------------------
// Getters — TurmaData
// ---------------------------------------------------------------------------

export function getTurma(chave) {
  return APP.turmas[chave] ?? null;
}

export function getTurmaAnterior(chave) {
  return APP.turmasAnterior[chave] ?? null;
}

export function getTurmas() {
  return APP.turmas;
}

export function hasTurmasAnterior() {
  return Object.keys(APP.turmasAnterior).length > 0;
}

export function getTurmaChaves() {
  return Object.keys(APP.turmas).sort();
}

export function hasTurmas() {
  return Object.keys(APP.turmas).length > 0;
}

// ---------------------------------------------------------------------------
// Getters — Digital
// ---------------------------------------------------------------------------

export function getDigital() {
  return APP.digital;
}

export function hasDigital() {
  return APP.digital.length > 0;
}

// ---------------------------------------------------------------------------
// Getters — Config
// ---------------------------------------------------------------------------

export function getConfig() {
  return APP.config;
}

// ---------------------------------------------------------------------------
// Getters — Derivados
// ---------------------------------------------------------------------------

export function getDisciplinaInfo(chaveTurma, nomeDisciplina) {
  const turma = APP.turmas[chaveTurma];
  if (!turma) return null;
  return turma.disciplinas.find(
    (d) => d.nome === nomeDisciplina || d.nomeOriginal === nomeDisciplina
  ) ?? null;
}

export function getAlunosAtivos(chaveTurma) {
  const turma = APP.turmas[chaveTurma];
  if (!turma) return [];
  return turma.alunos;
}

export function getTotalAlunosAtivos() {
  return Object.values(APP.turmas).reduce(
    (sum, turma) => sum + turma.alunos.length,
    0
  );
}

export function getTodasDisciplinas() {
  const set = new Set();
  for (const turma of Object.values(APP.turmas)) {
    for (const disc of turma.disciplinas) {
      set.add(disc.nome);
    }
  }
  return [...set].sort();
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

export function setTurma(chave, turmaData) {
  APP.turmas[chave] = turmaData;
}

export function setTurmaAnterior(chave, turmaData) {
  APP.turmasAnterior[chave] = turmaData;
}

export function setDigital(data) {
  APP.digital = data;
}

export function appendDigital(data) {
  APP.digital.push(...data);
}

export function setConfig(config) {
  APP.config = config;
}

export function clearData() {
  APP.turmas = {};
  APP.turmasAnterior = {};
  APP.digital = [];
}

export function removeTurma(chave) {
  delete APP.turmas[chave];
  delete APP.turmasAnterior[chave];
}
