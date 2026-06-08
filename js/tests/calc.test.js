/**
 * calc.test.js — Testes unitários das funções puras de calc.js.
 *
 * Executar: node --experimental-vm-modules js/tests/calc.test.js
 */

import assert from 'node:assert/strict';
import { calcFreq, resolverNota, calcMedia, calcSituacao, calcFreqMedia } from '../calc.js';

// ---------------------------------------------------------------------------
// Helper: mock mínimo de config
// ---------------------------------------------------------------------------

function makeConfig(overrides = {}) {
  return {
    params: {
      mediaAprov: 5.0,
      freqMin: 75,
      freqCrit: 50,
      ...(overrides.params || {}),
    },
    mencoes: {
      ET: 10,
      ES: 7,
      EP: 5,
      ...(overrides.mencoes || {}),
    },
  };
}

// ---------------------------------------------------------------------------
// T4 — Cálculo de Frequência (calcFreq)
// ---------------------------------------------------------------------------

{
  // T4.1: faltas=2, aulasDadas=20 → 90%
  assert.strictEqual(calcFreq(2, 20), 90, 'T4.1: 2 faltas em 20 aulas = 90%');

  // T4.2: faltas=0, aulasDadas=22 → 100%
  assert.strictEqual(calcFreq(0, 22), 100, 'T4.2: 0 faltas em 22 aulas = 100%');

  // T4.3: faltas=22, aulasDadas=22 → 0%
  assert.strictEqual(calcFreq(22, 22), 0, 'T4.3: 22 faltas em 22 aulas = 0%');

  // T4.4: aulasDadas=0 → null
  assert.strictEqual(calcFreq(0, 0), null, 'T4.4: aulasDadas=0 retorna null');
  assert.strictEqual(calcFreq(5, 0), null, 'T4.4b: aulasDadas=0 retorna null independente de faltas');

  // Valores decimais
  const freq = calcFreq(3, 10);
  assert.strictEqual(freq, 70, 'T4.5: 3 faltas em 10 aulas = 70%');
}

// ---------------------------------------------------------------------------
// T5 — Cálculo de Nota (resolverNota)
// ---------------------------------------------------------------------------

{
  const cfg = makeConfig();

  // T5.1: nota numérica → retorna o próprio número
  assert.strictEqual(resolverNota(7, cfg), 7);
  assert.strictEqual(resolverNota(5.5, cfg), 5.5);
  assert.strictEqual(resolverNota(0, cfg), 0);
  assert.strictEqual(resolverNota(10, cfg), 10);

  // T5.2: menção com equivalência configurada
  assert.strictEqual(resolverNota('ES', cfg), 7);
  assert.strictEqual(resolverNota('ET', cfg), 10);
  assert.strictEqual(resolverNota('EP', cfg), 5);

  // T5.3: menção SEM equivalência → null
  assert.strictEqual(resolverNota('MB', cfg), null);

  // T5.4: traço / hífen → null
  assert.strictEqual(resolverNota('-', cfg), null);
  assert.strictEqual(resolverNota('—', cfg), null);

  // T5.5: string vazia → null
  assert.strictEqual(resolverNota('', cfg), null);

  // T5.6: NaN numérico → null
  assert.strictEqual(resolverNota(NaN, cfg), null);

  // T5.7: sem config → null para menção
  assert.strictEqual(resolverNota('ES', null), null);
  assert.strictEqual(resolverNota(7, null), 7);
}

// ---------------------------------------------------------------------------
// T5 cont. — Cálculo de Média (calcMedia)
// ---------------------------------------------------------------------------

{
  const cfg = makeConfig();

  const discList = [
    { nome: 'MATEMATICA', aulasDadas: 20 },
    { nome: 'PORTUGUES', aulasDadas: 20 },
    { nome: 'ARTE', aulasDadas: 10 },
  ];

  // T5.8: notas numéricas em todas as disciplinas → média aritmética
  const aluno1 = { disciplinas: {
    MATEMATICA: { nota: 8, faltas: 0 },
    PORTUGUES: { nota: 6, faltas: 0 },
    ARTE: { nota: 7, faltas: 0 },
  }};
  assert.strictEqual(calcMedia(aluno1, discList, cfg), 7, 'T5.8: média de 8+6+7/3 = 7');

  // T5.9: menção com equivalência → entra no cálculo
  const aluno2 = { disciplinas: {
    MATEMATICA: { nota: 8, faltas: 0 },
    PORTUGUES: { nota: 'ES', faltas: 0 },  // ES=7
    ARTE: { nota: 6, faltas: 0 },
  }};
  assert.strictEqual(calcMedia(aluno2, discList, cfg), 7, 'T5.9: 8+7+6/3 = 7');

  // T5.10: menção sem equivalência → ignorada
  const aluno3 = { disciplinas: {
    MATEMATICA: { nota: 8, faltas: 0 },
    PORTUGUES: { nota: 'MB', faltas: 0 },  // sem equivalência
    ARTE: { nota: 6, faltas: 0 },
  }};
  assert.strictEqual(calcMedia(aluno3, discList, cfg), 7, 'T5.10: 8+6/2 = 7 (MB ignorado)');

  // T5.11: todas as disciplinas com menção sem equivalência → null
  const aluno4 = { disciplinas: {
    MATEMATICA: { nota: 'MB', faltas: 0 },
    PORTUGUES: { nota: 'B', faltas: 0 },
    ARTE: { nota: 'OT', faltas: 0 },
  }};
  assert.strictEqual(calcMedia(aluno4, discList, cfg), null, 'T5.11: todas sem equiv → null');

  // T5.12: traço em algumas disciplinas
  const aluno5 = { disciplinas: {
    MATEMATICA: { nota: 7, faltas: 0 },
    PORTUGUES: { nota: '-', faltas: 0 },
    ARTE: { nota: 9, faltas: 0 },
  }};
  assert.strictEqual(calcMedia(aluno5, discList, cfg), 8, 'T5.12: 7+9/2 = 8 (traço ignorado)');

  // T5.13: disciplina não está na discList → ignorada
  const aluno6 = { disciplinas: {
    MATEMATICA: { nota: 7, faltas: 0 },
  }};

  // ARTE should not be in discList for this test
  // Actually let me just use the existing discList but with aluno not having ARTE
  // That's what we have - aluno6 only has MATEMATICA
  assert.strictEqual(calcMedia(aluno6, discList, cfg), 7, 'T5.13: só MATEMATICA = 7');

  // T5.14: nota null (não string, null literal)
  const aluno7 = { disciplinas: {
    MATEMATICA: { nota: null, faltas: 0 },
    PORTUGUES: { nota: 8, faltas: 0 },
    ARTE: { nota: 6, faltas: 0 },
  }};
  assert.strictEqual(calcMedia(aluno7, discList, cfg), 7, 'T5.14: null ignorado, 8+6/2=7');

  // T5.15: uma única disciplina
  const discList2 = [{ nome: 'MATEMATICA', aulasDadas: 20 }];
  const aluno8 = { disciplinas: { MATEMATICA: { nota: 9.5, faltas: 0 } } };
  assert.strictEqual(calcMedia(aluno8, discList2, cfg), 9.5, 'T5.15: disciplina única = 9.5');
}

// ---------------------------------------------------------------------------
// T6 — Situação do Aluno (calcSituacao)
// ---------------------------------------------------------------------------

{
  const cfg = makeConfig();

  const turma = {
    disciplinas: [
      { nome: 'MATEMATICA', aulasDadas: 20 },
      { nome: 'PORTUGUES', aulasDadas: 20 },
      { nome: 'ARTE', aulasDadas: 10 },
    ],
  };

  function makeAluno(notas, faltas) {
    const disc = {};
    for (const d of turma.disciplinas) {
      disc[d.nome] = {
        nota: notas[d.nome] ?? null,
        faltas: faltas[d.nome] ?? 0,
      };
    }
    return { disciplinas: disc };
  }

  // T6.1: Todas as notas ≥ 5,0; todas as freq. ≥ 75% → Aprovado
  const a1 = makeAluno({ MATEMATICA: 7, PORTUGUES: 8, ARTE: 6 }, { MATEMATICA: 0, PORTUGUES: 2, ARTE: 0 });
  {
    const { sit, nivelFreq } = calcSituacao(a1, turma, cfg);
    assert.strictEqual(sit, 'Aprovado', 'T6.1: Aprovado com boas notas e freq');
    assert.strictEqual(nivelFreq, null, 'T6.1: sem alerta de freq');
  }

  // T6.2: Qualquer nota < 5,0 → Reprovado
  const a2 = makeAluno({ MATEMATICA: 4, PORTUGUES: 8, ARTE: 7 }, { MATEMATICA: 0, PORTUGUES: 0, ARTE: 0 });
  {
    const { sit } = calcSituacao(a2, turma, cfg);
    assert.strictEqual(sit, 'Reprovado', 'T6.2: Reprovado com nota < 5');
  }

  // T6.3: Aprovado por nota; freq entre 50% e 74% → alerta
  const a3 = makeAluno({ MATEMATICA: 7, PORTUGUES: 8, ARTE: 6 }, { MATEMATICA: 6, PORTUGUES: 5, ARTE: 2 });
  // MAT: 14/20=70%, PORT: 15/20=75%, ARTE: 8/10=80% → min=70% → alerta
  {
    const { sit, nivelFreq } = calcSituacao(a3, turma, cfg);
    assert.strictEqual(sit, 'Aprovado', 'T6.3: Aprovado por nota');
    assert.strictEqual(nivelFreq, 'alerta', 'T6.3: freq alerta (70% < 75%)');
  }

  // T6.4: Aprovado por nota; freq < 50% → crítico
  const a4 = makeAluno({ MATEMATICA: 8, PORTUGUES: 8, ARTE: 7 }, { MATEMATICA: 15, PORTUGUES: 5, ARTE: 2 });
  // MAT: 5/20=25% → crítico
  {
    const { sit, nivelFreq } = calcSituacao(a4, turma, cfg);
    assert.strictEqual(sit, 'Aprovado', 'T6.4: Aprovado por nota');
    assert.strictEqual(nivelFreq, 'critico', 'T6.4: freq crítica (25% < 50%)');
  }

  // T6.5: Reprovado por nota com freq crítica → Reprovado (nível exibido à parte)
  const a5 = makeAluno({ MATEMATICA: 3, PORTUGUES: 5, ARTE: 4 }, { MATEMATICA: 15, PORTUGUES: 5, ARTE: 8 });
  {
    const { sit, nivelFreq } = calcSituacao(a5, turma, cfg);
    assert.strictEqual(sit, 'Reprovado', 'T6.5: Reprovado por nota');
    // ainda retorna o nível de freq para exibição
    assert.strictEqual(nivelFreq, 'critico', 'T6.5: freq crítica detectada');
  }

  // T6.6: Menção "ES" com equivalência 7 → nota < 5? Não (7 ≥ 5), então Aprovado
  const a6 = makeAluno({ MATEMATICA: 'ES', PORTUGUES: 8, ARTE: 6 }, { MATEMATICA: 0, PORTUGUES: 0, ARTE: 0 });
  {
    const { sit } = calcSituacao(a6, turma, cfg);
    assert.strictEqual(sit, 'Aprovado', 'T6.6: ES=7 ≥ 5 → Aprovado');
  }

  // T6.7: Sem equivalência para menção → menção ignorada; reprova nas outras
  const a7 = makeAluno({ MATEMATICA: 'MB', PORTUGUES: 3, ARTE: 4 }, { MATEMATICA: 0, PORTUGUES: 0, ARTE: 0 });
  {
    const { sit } = calcSituacao(a7, turma, cfg);
    assert.strictEqual(sit, 'Reprovado', 'T6.7: MB ignorada, PORT 3 < 5 → Reprovado');
  }

  // T6.8: Todas as menções sem equivalência → sem nota resolvida → Aprovado
  const a8 = makeAluno({ MATEMATICA: 'MB', PORTUGUES: 'B', ARTE: 'OT' }, { MATEMATICA: 0, PORTUGUES: 0, ARTE: 0 });
  {
    const { sit } = calcSituacao(a8, turma, cfg);
    assert.strictEqual(sit, 'Aprovado', 'T6.8: sem notas para julgar → Aprovado');
  }

  // T6.9: aulasDadas=0 → freq retorna null → não afeta nível de freq
  const turmaComZero = {
    disciplinas: [
      { nome: 'MATEMATICA', aulasDadas: 0 },
      { nome: 'PORTUGUES', aulasDadas: 20 },
    ],
  };
  const a9 = makeAluno({ MATEMATICA: 7, PORTUGUES: 8 }, { MATEMATICA: 5, PORTUGUES: 0 });
  {
    const { sit, nivelFreq } = calcSituacao(a9, turmaComZero, cfg);
    assert.strictEqual(sit, 'Aprovado', 'T6.9: Aprovado');
    assert.strictEqual(nivelFreq, null, 'T6.9: freq de aulasDadas=0 retorna null, PORT=100% → sem alerta');
  }

  // T6.10: Exatamente 75% de frequência → NÃO é alerta (≥ freqMin)
  const a10 = makeAluno({ MATEMATICA: 7, PORTUGUES: 8, ARTE: 7 }, { MATEMATICA: 5, PORTUGUES: 0, ARTE: 0 });
  // MAT: 15/20=75% → NOT alerta, PORT: 20/20=100%, ARTE: 10/10=100% → min=75%
  {
    const { nivelFreq } = calcSituacao(a10, turma, cfg);
    assert.strictEqual(nivelFreq, null, 'T6.10: 75% ≥ 75 → sem alerta');
  }

  // T6.11: Exatamente 50% → crítico (< freqCrit que é 50? NÃO, 50 ≥ 50)
  // Wait, freqCrit = 50, so < 50 is critico. 50% is NOT critico, it's alerta (50 < 75)
  const a11 = makeAluno({ MATEMATICA: 7, PORTUGUES: 8, ARTE: 7 }, { MATEMATICA: 10, PORTUGUES: 0, ARTE: 0 });
  // MAT: 10/20=50%, min=50%
  {
    const { nivelFreq } = calcSituacao(a11, turma, cfg);
    assert.strictEqual(nivelFreq, 'alerta', 'T6.11: 50% ≥ freqCrit → alerta, não crítico');
  }

  // T6.12: Exatamente 49% (arredondado): 9.8/20 = 49% → freq < 50 → crítico
  // Let me use clean numbers: 11 faltas em 20 aulas = 9/20 = 45%
  const a12 = makeAluno({ MATEMATICA: 7, PORTUGUES: 8, ARTE: 7 }, { MATEMATICA: 11, PORTUGUES: 0, ARTE: 0 });
  // MAT: 9/20=45%
  {
    const { nivelFreq } = calcSituacao(a12, turma, cfg);
    assert.strictEqual(nivelFreq, 'critico', 'T6.12: 45% < 50 → crítico');
  }
}

// ---------------------------------------------------------------------------
// calcFreqMedia — Frequência média do aluno
// ---------------------------------------------------------------------------

{
  const discList = [
    { nome: 'MATEMATICA', aulasDadas: 20 },
    { nome: 'PORTUGUES', aulasDadas: 20 },
    { nome: 'ARTE', aulasDadas: 10 },
  ];

  const aluno = { disciplinas: {
    MATEMATICA: { nota: 7, faltas: 2 },
    PORTUGUES: { nota: 8, faltas: 5 },
    ARTE: { nota: 6, faltas: 1 },
  }};
  // MAT: 18/20=90, PORT: 15/20=75, ARTE: 9/10=90
  // total: (18+15+9)/(20+20+10) = 42/50 = 84%
  assert.strictEqual(calcFreqMedia(aluno, discList), 84, 'Freq média: 84%');

  // aulasDadas=0 → disciplina ignorada
  const discList2 = [
    { nome: 'MATEMATICA', aulasDadas: 0 },
    { nome: 'PORTUGUES', aulasDadas: 20 },
  ];
  const aluno2 = { disciplinas: {
    MATEMATICA: { nota: 7, faltas: 5 },
    PORTUGUES: { nota: 8, faltas: 0 },
  }};
  assert.strictEqual(calcFreqMedia(aluno2, discList2), 100, 'Freq média ignora aulasDadas=0');

  // Sem disciplinas → null
  assert.strictEqual(calcFreqMedia(aluno, []), null, 'Freq média sem disciplinas = null');
}

console.log('✅ calc.test.js — Todos os testes passaram!');
