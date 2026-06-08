/**
 * parser.js — Leitura e normalização de arquivos xlsx (Mapão + Digital).
 *
 * SheetJS (XLSX) deve estar disponível como global (carregado via CDN).
 */

// Referência defensiva ao XLSX global (CDN script tag)
const _XLSX = typeof XLSX !== 'undefined' ? XLSX : globalThis.XLSX;

// ---------------------------------------------------------------------------
// Normalização de nomes
// ---------------------------------------------------------------------------

/**
 * Remove acentos e caracteres especiais, retorna uppercase.
 */
export function normalizarDisc(raw) {
  return raw
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9 ]/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

/**
 * Normaliza nome da turma para chave canônica: "1A", "2B", etc.
 * Fallback: uppercase, strip spaces, first 4 chars.
 */
export function normalizarTurma(str) {
  if (!str) return '';

  // Padrão principal: "1ª Série A" ou "1ª SERIE A INTEGRAL ..."
  const match1 = str.match(/(\d)[ªº°]?\s*[Ss][ÉEée][Rr][Ii][Ee]?\s*([A-Za-z])/);
  if (match1) {
    return match1[1] + match1[2].toUpperCase();
  }

  // Padrão alternativo: "1ª A" (sem "Série")
  const match2 = str.match(/(\d)\s*[ªº°]\s*([A-Za-z])/);
  if (match2) {
    return match2[1] + match2[2].toUpperCase();
  }

  // Padrão: "1A" direto
  const match3 = str.match(/^(\d)([A-Za-z])$/);
  if (match3) {
    return match3[1] + match3[2].toUpperCase();
  }

  // Fallback
  return str.toUpperCase().replace(/\s/g, '').slice(0, 4);
}

// ---------------------------------------------------------------------------
// parseMapao — Lê um arquivo de Mapão e retorna TurmaData
// ---------------------------------------------------------------------------

/**
 * @param {File|ArrayBuffer|Uint8Array} file — arquivo xlsx
 * @returns {object} TurmaData
 */
export function parseMapao(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = _XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];

        // Usar header:0 / aoa — SEM header:1
        const aoa = _XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        const turmaData = _parseAoa(aoa);
        resolve(turmaData);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('Erro ao ler arquivo'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Versão síncrona para uso em Node.js (testes).
 * @param {Array[]} aoa — Array of arrays do SheetJS
 */
export function _parseAoa(aoa) {
  // Extrair nomeOriginal: linha 5 (index 5), coluna 1 (index 1)
  const nomeOriginal = (aoa[5] && aoa[5][1]) ? String(aoa[5][1]).replace('Turma:', '').trim() : '';

  // Extrair bimestre: linha 6 (index 6), coluna 1
  let bimestre = 0;
  if (aoa[6] && aoa[6][1]) {
    const tipoStr = String(aoa[6][1]).toLowerCase();
    // Formatos: "Conselho Primeiro Bimestre" ou "1º Bimestre"
    if (tipoStr.includes('primeiro') || tipoStr.includes('1º')) bimestre = 1;
    else if (tipoStr.includes('segundo') || tipoStr.includes('2º')) bimestre = 2;
    else if (tipoStr.includes('terceiro') || tipoStr.includes('3º')) bimestre = 3;
    else if (tipoStr.includes('quarto') || tipoStr.includes('4º')) bimestre = 4;
    // Fallback: regex numérico
    if (bimestre === 0) {
      const bimMatch = tipoStr.match(/(\d+)º/);
      if (bimMatch) bimestre = parseInt(bimMatch[1], 10);
    }
  }

  // Chave canônica
  const chave = normalizarTurma(nomeOriginal);

  // Detectar disciplinas: linha 10 (index 10), de col=2 step=4
  const linhaHeader = aoa[10] || [];
  const ncols = linhaHeader.length;
  const disciplinas = [];

  for (let col = 2; col < ncols; col += 4) {
    const cell = linhaHeader[col];
    if (cell === null || cell === undefined) continue;
    const cellStr = String(cell).trim();
    if (cellStr === 'TOTAL') break;

    const nomeCompleto = cellStr.split('\n')[0].trim();
    const nomeNormalizado = normalizarDisc(nomeCompleto);

    disciplinas.push({
      nome: nomeNormalizado,
      nomeOriginal: nomeCompleto,
      aulasDadas: 0,
      colIndex: col,
    });
  }

  // Detectar linha do rodapé (Aulas Dadas)
  let rodapeIdx = -1;
  for (let r = aoa.length - 1; r >= 0; r--) {
    const row = aoa[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell && typeof cell === 'string' && cell.trim().startsWith('Aulas Dadas:')) {
        rodapeIdx = r;
        break;
      }
    }
    if (rodapeIdx !== -1) break;
  }

  // Extrair aulasDadas do rodapé
  if (rodapeIdx !== -1) {
    const rodapeRow = aoa[rodapeIdx];
    for (const disc of disciplinas) {
      const val = rodapeRow[disc.colIndex];
      if (val !== null && val !== undefined) {
        // Célula contém "Aulas Dadas: 22" — extrair número
        const numMatch = String(val).match(/(\d+)/);
        disc.aulasDadas = numMatch ? parseInt(numMatch[1], 10) : 0;
      }
    }
  }

  // Detectar alunos: linha 12 (index 12) até rodapeIdx - 1
  const alunos = [];
  const startRow = 12;
  const endRow = rodapeIdx !== -1 ? rodapeIdx - 1 : aoa.length - 1;

  for (let r = startRow; r <= endRow; r++) {
    const row = aoa[r];
    if (!row) continue;

    // Coluna 1 (index 1) = SITUAÇÃO
    const situacaoRaw = row[1];
    const situacao = situacaoRaw ? String(situacaoRaw).trim() : '';

    // Filtrar apenas "Ativo"
    if (situacao !== 'Ativo') continue;

    // Coluna 0 (index 0) = nome do aluno
    const nome = row[0] ? String(row[0]).trim() : '';

    // Extrair notas por disciplina
    const alunoDisciplinas = {};
    for (const disc of disciplinas) {
      const notaRaw = row[disc.colIndex + 1];
      const faltasRaw = row[disc.colIndex + 2];
      const acRaw = row[disc.colIndex + 3];

      let nota = null;
      if (notaRaw !== null && notaRaw !== undefined && notaRaw !== '') {
        const notaStr = String(notaRaw).trim();
        if (notaStr === '-' || notaStr === '—') {
          nota = null;
        } else {
          const parsed = parseFloat(notaStr.replace(',', '.'));
          if (!isNaN(parsed)) {
            nota = parsed;
          } else {
            nota = notaStr; // menção (string)
          }
        }
      }

      const faltas = faltasRaw !== null && faltasRaw !== undefined
        ? (parseInt(faltasRaw, 10) || 0)
        : 0;

      const ac = acRaw !== null && acRaw !== undefined
        ? (parseInt(acRaw, 10) || 0)
        : 0;

      alunoDisciplinas[disc.nome] = { nota, faltas, ac };
    }

    alunos.push({
      nome,
      disciplinas: alunoDisciplinas,
    });
  }

  return {
    nomeOriginal,
    chave,
    bimestre,
    disciplinas,
    alunos,
  };
}

// ---------------------------------------------------------------------------
// parseDigital — Lê a planilha de aulas digitais
// ---------------------------------------------------------------------------

/**
 * @param {File|ArrayBuffer|Uint8Array} file — arquivo xlsx
 * @returns {Array} [DigitalRow, ...]
 */
export function parseDigital(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = _XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const aoa = _XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        const result = _parseDigitalAoa(aoa);
        resolve(result);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('Erro ao ler arquivo'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * @param {Array[]} aoa
 * @returns {Array}
 */
export function _parseDigitalAoa(aoa) {
  if (aoa.length < 2) return [];

  const header = aoa[0];
  const colIndex = {};

  for (let i = 0; i < header.length; i++) {
    const h = header[i] ? String(header[i]).trim().toUpperCase() : '';
    if (h.includes('TURMA')) colIndex.turma = i;
    if (h.includes('DISCIPLINA')) colIndex.disciplina = i;
    if (h.includes('PREVISTO') && h.includes('1')) colIndex.prev1 = i;
    if (h.includes('CONCLUIDO') && h.includes('1')) colIndex.conc1 = i;
    if (h.includes('PREVISTO') && h.includes('2')) colIndex.prev2 = i;
    if (h.includes('CONCLUIDO') && h.includes('2')) colIndex.conc2 = i;
    if (h.includes('PREVISTO') && h.includes('3')) colIndex.prev3 = i;
    if (h.includes('CONCLUIDO') && h.includes('3')) colIndex.conc3 = i;
    if (h.includes('PREVISTO') && h.includes('4')) colIndex.prev4 = i;
    if (h.includes('CONCLUIDO') && h.includes('4')) colIndex.conc4 = i;
  }

  const result = [];

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || !row[colIndex.turma]) continue;

    const chave = normalizarTurma(String(row[colIndex.turma]).trim());
    const disciplina = normalizarDisc(String(row[colIndex.disciplina] || '').trim());

    const bimestres = [
      {
        previsto: _parseIntNull(row[colIndex.prev1]),
        concluido: _parseIntNull(row[colIndex.conc1]),
      },
      {
        previsto: _parseIntNull(row[colIndex.prev2]),
        concluido: _parseIntNull(row[colIndex.conc2]),
      },
      {
        previsto: _parseIntNull(row[colIndex.prev3]),
        concluido: _parseIntNull(row[colIndex.conc3]),
      },
      {
        previsto: _parseIntNull(row[colIndex.prev4]),
        concluido: _parseIntNull(row[colIndex.conc4]),
      },
    ];

    result.push({ chave, disciplina, bimestres });
  }

  return result;
}

function _parseIntNull(val) {
  if (val === null || val === undefined || val === '') return null;
  const parsed = parseInt(val, 10);
  return isNaN(parsed) ? null : parsed;
}
