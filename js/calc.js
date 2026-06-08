/**
 * calc.js — Funções puras de cálculo.
 *
 * NENHUMA dependência de DOM, store, ou módulos de browser.
 * Pode ser importado por qualquer módulo (inclusive testes Node.js).
 */

// ---------------------------------------------------------------------------
// calcFreq — Frequência por aluno por disciplina
// ---------------------------------------------------------------------------

/**
 * @param {number} faltas
 * @param {number} aulasDadas
 * @returns {number|null} percentual de frequência, ou null se aulasDadas === 0
 */
export function calcFreq(faltas, aulasDadas) {
  if (aulasDadas === 0) return null;
  return ((aulasDadas - faltas) / aulasDadas) * 100;
}

// ---------------------------------------------------------------------------
// resolverNota — Converte menção em número (se houver equivalência)
// ---------------------------------------------------------------------------

/**
 * @param {number|string} nota — número ou string de menção
 * @param {object} config — APP.config
 * @returns {number|null} nota numérica, ou null se menção sem equivalência
 */
export function resolverNota(nota, config) {
  if (typeof nota === 'number' && !Number.isNaN(nota)) {
    return nota;
  }

  if (typeof nota === 'string') {
    const trimmed = nota.trim();
    if (trimmed === '-' || trimmed === '—' || trimmed === 'null' || trimmed === '') {
      return null;
    }
    if (config && config.mencoes && trimmed in config.mencoes) {
      const equiv = config.mencoes[trimmed];
      if (typeof equiv === 'number' && !Number.isNaN(equiv)) {
        return equiv;
      }
    }
    return null;
  }

  return null;
}

// ---------------------------------------------------------------------------
// calcMedia — Média aritmética do aluno
// ---------------------------------------------------------------------------

/**
 * Calcula a média aritmética das notas resolvidas (não-nulas).
 * Se nenhuma disciplina tem nota resolvida, retorna null.
 *
 * @param {object} alunoData — { disciplinas: { "MATEMATICA": { nota, ... }, ... } }
 * @param {Array} discList   — array de DisciplinaInfo da turma
 * @param {object} config    — APP.config
 * @returns {number|null}
 */
export function calcMedia(alunoData, discList, config) {
  let soma = 0;
  let count = 0;

  for (const disc of discList) {
    const dados = alunoData.disciplinas[disc.nome];
    if (!dados) continue;

    const nota = resolverNota(dados.nota, config);
    if (nota !== null) {
      soma += nota;
      count++;
    }
  }

  return count > 0 ? soma / count : null;
}

// ---------------------------------------------------------------------------
// calcSituacao — Situação do aluno (Aprovado/Reprovado + nível de frequência)
// ---------------------------------------------------------------------------

/**
 * @param {object} alunoData
 * @param {object} turmaData  — contém disciplinas com aulasDadas
 * @param {object} config     — APP.config
 * @returns {{ sit: string, nivelFreq: string|null }}
 */
export function calcSituacao(alunoData, turmaData, config) {
  const mediaAprov = config.params.mediaAprov;
  const freqMin = config.params.freqMin;
  const freqCrit = config.params.freqCrit;

  let reprovado = false;
  let temNota = false;
  let minFreq = null;

  for (const disc of turmaData.disciplinas) {
    const dados = alunoData.disciplinas[disc.nome];
    if (!dados) continue;

    // Verificar nota
    const nota = resolverNota(dados.nota, config);
    if (nota !== null) {
      temNota = true;
      if (nota < mediaAprov) {
        reprovado = true;
      }
    }

    // Verificar frequência
    const freq = calcFreq(dados.faltas, disc.aulasDadas);
    if (freq !== null) {
      if (minFreq === null || freq < minFreq) {
        minFreq = freq;
      }
    }
  }

  // Situação
  const sit = reprovado ? 'Reprovado' : (temNota || minFreq !== null ? 'Aprovado' : 'Aprovado');

  // Nível de frequência (apenas para Aprovados — Reprovados mostram alerta separado)
  let nivelFreq = null;
  if (minFreq !== null) {
    if (minFreq < freqCrit) {
      nivelFreq = 'critico';
    } else if (minFreq < freqMin) {
      nivelFreq = 'alerta';
    }
  }

  return { sit, nivelFreq };
}

// ---------------------------------------------------------------------------
// calcFreqMedia — Frequência média considerando todas as disciplinas do aluno
// ---------------------------------------------------------------------------

/**
 * Calcula a frequência média do aluno considerando todas as disciplinas.
 * Soma (aulasDadas - faltas) de todas as disciplinas / soma aulasDadas.
 *
 * @param {object} alunoData
 * @param {Array} discList
 * @returns {number|null}
 */
export function calcFreqMedia(alunoData, discList) {
  let totalAulas = 0;
  let totalPresente = 0;

  for (const disc of discList) {
    const dados = alunoData.disciplinas[disc.nome];
    if (!dados) continue;
    if (disc.aulasDadas === 0) continue;

    totalAulas += disc.aulasDadas;
    totalPresente += (disc.aulasDadas - dados.faltas);
  }

  return totalAulas > 0 ? (totalPresente / totalAulas) * 100 : null;
}
