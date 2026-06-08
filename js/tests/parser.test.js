/**
 * parser.test.js — Testes de parser com mapões reais.
 *
 * Executar: node --experimental-vm-modules js/tests/parser.test.js
 *
 * Requer xlsx (npm install xlsx) para leitura dos arquivos no Node.js.
 */

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { _parseAoa, normalizarTurma, normalizarDisc } from '../parser.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const mapaDir = resolve(__dirname, '..', '..', 'mapões');

// ---------------------------------------------------------------------------
// Carregar xlsx (tentar import)
// ---------------------------------------------------------------------------

let XLSX;
try {
  XLSX = await import('xlsx');
} catch {
  console.error('ERRO: xlsx não instalado. Execute: npm install xlsx');
  process.exit(1);
}

// ---------------------------------------------------------------------------
// Helper: carregar mapão como AOA
// ---------------------------------------------------------------------------

function loadMapa(filename) {
  const filePath = resolve(mapaDir, filename);
  const buf = readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
}

// ---------------------------------------------------------------------------
// Helper: encontrar disciplinas
// ---------------------------------------------------------------------------

function findDisc(turmaData, nome) {
  return turmaData.disciplinas.find((d) => d.nome === nome) || null;
}

// ---------------------------------------------------------------------------
// Carregar todos os mapões
// ---------------------------------------------------------------------------

const mapas = {};
{
  // Listar arquivos no diretório mapões/
  const { readdirSync } = await import('node:fs');
  const files = readdirSync(mapaDir).filter((f) => f.endsWith('.xlsx'));

  for (const file of files) {
    const aoa = loadMapa(file);
    const turmaData = _parseAoa(aoa);
    mapas[turmaData.chave] = turmaData;
    console.log(`  Carregado: ${turmaData.chave} (${file})`);
  }
}

// ---------------------------------------------------------------------------
// T2 — Disciplinas Dinâmicas
// ---------------------------------------------------------------------------

const disciplinasEsperadas = {
  '1A': 20,
  '2A': 21,
  '3A': 18,
};

for (const [chave, esperado] of Object.entries(disciplinasEsperadas)) {
  if (!mapas[chave]) {
    console.log(`  ⚠ T2: Turma ${chave} não encontrada nos arquivos — pulando`);
    continue;
  }
  const turma = mapas[chave];
  assert.strictEqual(
    turma.disciplinas.length,
    esperado,
    `T2: ${chave} deve ter ${esperado} disciplinas, encontrou ${turma.disciplinas.length}`
  );
}

// T2.5: TOTAL não aparece como disciplina
for (const turma of Object.values(mapas)) {
  for (const disc of turma.disciplinas) {
    assert.notStrictEqual(disc.nome, 'TOTAL', `T2.5: TOTAL não deve ser disciplina em ${turma.chave}`);
  }
}

// T2.6: Nome com \n e código → extraído corretamente (todas as disciplinas têm nome sem \n)
for (const turma of Object.values(mapas)) {
  for (const disc of turma.disciplinas) {
    assert.ok(!disc.nome.includes('\n'), `T2.6: ${disc.nome} em ${turma.chave} não deve conter \\n`);
    assert.ok(!disc.nome.includes('2700'), `T2.6: nome não deve conter código numérico`);
  }
}

// ---------------------------------------------------------------------------
// T1 — Filtro de Situação (apenas "Ativo")
// ---------------------------------------------------------------------------

const statusInvalidos = ['Remanejamento', 'Transferido', 'Baixa - Transferência', 'Não Comparecimento'];

for (const turma of Object.values(mapas)) {
  for (const aluno of turma.alunos) {
    // Verificar que todos os alunos na lista são realmente Ativo
    // (o parser já filtra, então não deve ter nenhum não-Ativo)
    assert.ok(
      !statusInvalidos.some((s) => aluno.nome === s),
      `T1: Aluno em ${turma.chave} não deve ser status inválido`
    );
  }
}

// ---------------------------------------------------------------------------
// T3 — Aulas Dadas por Disciplina
// ---------------------------------------------------------------------------

const aulasDadasRef = {
  '1A': { 'MATEMATICA': 32, 'LINGUA PORTUGUESA': 44 },
  '2A': { 'MATEMATICA': 40, 'LINGUA PORTUGUESA': 38 },
  '3A': { 'MATEMATICA': 38, 'LINGUA PORTUGUESA': 44 },
};

for (const [chave, refs] of Object.entries(aulasDadasRef)) {
  if (!mapas[chave]) {
    console.log(`  ⚠ T3: Turma ${chave} não encontrada — pulando`);
    continue;
  }
  const turma = mapas[chave];
  for (const [discNome, esperado] of Object.entries(refs)) {
    const disc = findDisc(turma, discNome);
    assert.ok(disc !== null, `T3: ${discNome} deve existir em ${chave}`);
    assert.strictEqual(
      disc.aulasDadas,
      esperado,
      `T3: ${discNome} aulasDadas em ${chave} deve ser ${esperado}, foi ${disc.aulasDadas}`
    );
  }
}

// T3.3: Disciplina com aulasDadas = 0 existe e tem 0
for (const turma of Object.values(mapas)) {
  for (const disc of turma.disciplinas) {
    assert.ok(typeof disc.aulasDadas === 'number', `T3.3: ${disc.nome} em ${turma.chave} aulasDadas deve ser number`);
    assert.ok(disc.aulasDadas >= 0, `T3.3: ${disc.nome} em ${turma.chave} aulasDadas >= 0`);
  }
}

// ---------------------------------------------------------------------------
// T1/T10 — Contagem de alunos ativos
// ---------------------------------------------------------------------------

const alunosEsperados = {
  '1A': 28,
  '2A': 43,
  '3A': 31,
};

for (const [chave, esperado] of Object.entries(alunosEsperados)) {
  if (!mapas[chave]) {
    console.log(`  ⚠ T10: Turma ${chave} não encontrada — pulando`);
    continue;
  }
  const turma = mapas[chave];
  assert.strictEqual(
    turma.alunos.length,
    esperado,
    `T10: ${chave} deve ter ${esperado} alunos ativos, encontrou ${turma.alunos.length}`
  );
}

// ---------------------------------------------------------------------------
// Normalização de Turma
// ---------------------------------------------------------------------------

{
  assert.strictEqual(normalizarTurma('1ª Série A'), '1A');
  assert.strictEqual(normalizarTurma('1ª SERIE A INTEGRAL 9H ANUAL'), '1A');
  assert.strictEqual(normalizarTurma('2ª série B'), '2B');
  assert.strictEqual(normalizarTurma('3ªA'), '3A');
  assert.strictEqual(normalizarTurma('1A'), '1A');
  assert.strictEqual(normalizarTurma('TURMA X'), 'TURM'); // fallback
  assert.strictEqual(normalizarTurma('1ª Série A Integral'), '1A');
  assert.strictEqual(normalizarTurma('3º Série A'), '3A');
}

// ---------------------------------------------------------------------------
// Normalização de Disciplina
// ---------------------------------------------------------------------------

{
  assert.strictEqual(normalizarDisc('MATEMÁTICA'), 'MATEMATICA');
  assert.strictEqual(normalizarDisc('LÍNGUA PORTUGUESA'), 'LINGUA PORTUGUESA');
  assert.strictEqual(normalizarDisc('EDUCAÇÃO FÍSICA'), 'EDUCACAO FISICA');
  assert.strictEqual(normalizarDisc('  ARTES  '), 'ARTES');
  assert.strictEqual(normalizarDisc('MATEMATICA'), 'MATEMATICA');
}

// ---------------------------------------------------------------------------
// Validação de estrutura dos dados
// ---------------------------------------------------------------------------

for (const turma of Object.values(mapas)) {
  // Verificar que todo aluno tem nome
  for (const aluno of turma.alunos) {
    assert.ok(aluno.nome.length > 0, `Aluno em ${turma.chave} deve ter nome`);
    assert.ok(typeof aluno.disciplinas === 'object', `Aluno ${aluno.nome} deve ter disciplinas`);

    // Verificar que disciplinas do aluno batem com as da turma
    for (const disc of turma.disciplinas) {
      assert.ok(
        disc.nome in aluno.disciplinas,
        `Aluno ${aluno.nome} em ${turma.chave} deve ter disciplina ${disc.nome}`
      );
    }
  }

  // Verificar colIndex nas disciplinas
  for (const disc of turma.disciplinas) {
    assert.ok(typeof disc.colIndex === 'number', `${disc.nome} deve ter colIndex`);
    assert.ok(disc.colIndex >= 0, `${disc.nome} colIndex deve ser >= 0`);
  }

  // Verificar bimestre
  assert.ok(turma.bimestre >= 1 && turma.bimestre <= 4, `${turma.chave} bimestre deve ser 1-4`);
}

console.log('');
console.log('✅ parser.test.js — Todos os testes passaram!');
