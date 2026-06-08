/**
 * pptx.js — Geração de apresentações PPTX estilo Conselho de Classe.
 *
 * Dois modos:
 *   - "conselho": completo (destaques + alertas + encaminhamentos)
 *   - "pais":     somente positivo (destaques, sem alertas)
 */

import { getTurma, getTurmaChaves, hasTurmas, getConfig } from './store.js';
import { calcSituacao, calcMedia, calcFreqMedia, resolverNota, calcFreq } from './calc.js';

const _Pptx = typeof PptxGenJS !== 'undefined' ? PptxGenJS : globalThis.PptxGenJS;

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

export function initPPTX() {
  const btnConselho = document.getElementById('btn-gerar-pptx-conselho');
  const btnPais = document.getElementById('btn-gerar-pptx-pais');
  const selectTurma = document.getElementById('pptx-turma');
  const tipoSelect = document.getElementById('pptx-tipo');
  const discGroup = document.getElementById('pptx-disciplina-group');

  // Mostrar/ocultar disciplina
  if (tipoSelect && discGroup) {
    tipoSelect.addEventListener('change', () => {
      discGroup.style.display = tipoSelect.value === 'disciplina' ? '' : 'none';
    });
  }

  const getFiltroTurma = () => selectTurma?.value || 'todas';

  if (btnConselho) btnConselho.addEventListener('click', () => gerarPPTX('conselho', getFiltroTurma()));
  if (btnPais) btnPais.addEventListener('click', () => gerarPPTX('pais', getFiltroTurma()));

  // Populate turma selector
  window.addEventListener('cc:data-changed', () => {
    if (selectTurma && hasTurmas()) {
      const current = selectTurma.value || 'todas';
      const chaves = getTurmaChaves();
      selectTurma.innerHTML = '<option value="todas">Todas as turmas</option>' +
        chaves.map((c) => `<option value="${c}" ${c === current ? 'selected' : ''}>Turma ${c}</option>`).join('');
    }
  });
}

// ---------------------------------------------------------------------------
// Gerar PPTX
// ---------------------------------------------------------------------------

function gerarPPTX(modo, filtroTurma = 'todas') {
  if (!hasTurmas()) {
    alert('Importe os mapões antes de gerar apresentações.');
    return;
  }

  const pptx = new _Pptx();
  const config = getConfig();
  const todasChaves = getTurmaChaves();
  const chaves = filtroTurma === 'todas' ? todasChaves : [filtroTurma];

  const { destaques, tierOuro, tierPrata, tierBronze, freqPerfeita, alertas, dadosArea } =
    coletarDados(chaves, config);

  // 1. Capa
  criarCapa(pptx, config, chaves);

  // 2-4. Destaques por turma (notas + frequência)
  for (const chave of chaves) {
    criarSlideDestaquesTurma(pptx, chave, config);
  }

  // 5. Destaque de Frequência (100%)
  if (freqPerfeita.length > 0) {
    criarSlideFreqPerfeita(pptx, freqPerfeita);
  }

  // 6-8. Ouro / Prata / Bronze
  if (tierOuro.length > 0) criarSlideTier(pptx, '🥇 Destaque Ouro', tierOuro, 'Freq ≥ 90% · Média ≥ 9', 'C4953A');
  if (tierPrata.length > 0) criarSlideTier(pptx, '🥈 Destaque Prata', tierPrata, 'Freq ≥ 90% · Média 8–9', '8B95A3');
  if (tierBronze.length > 0) criarSlideTier(pptx, '🥉 Destaque Bronze', tierBronze, 'Freq ≥ 90% · Média 7–8', 'C26547');

  // 9+. Por Área
  for (const [areaNome, dados] of Object.entries(dadosArea)) {
    criarSlideArea(pptx, areaNome, dados, config);
  }

  // X. Encaminhamentos (somente modo conselho)
  if (modo === 'conselho' && alertas.length > 0) {
    criarSlideAlertas(pptx, alertas, config);
  }

  const filename = modo === 'conselho'
    ? 'apresentacao-conselho.pptx'
    : 'apresentacao-pais.pptx';

  pptx.writeFile({ fileName: filename });
}

// ---------------------------------------------------------------------------
// Coleta de dados
// ---------------------------------------------------------------------------

function coletarDados(chaves, config) {
  const destaques = [];     // top 5 por turma
  const tierOuro = [];
  const tierPrata = [];
  const tierBronze = [];
  const freqPerfeita = [];
  const alertas = [];
  const dadosArea = {};

  for (const chave of chaves) {
    const turma = getTurma(chave);
    const turmaAlunos = [];

    for (const aluno of turma.alunos) {
      const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
      const media = calcMedia(aluno, turma.disciplinas, config);
      const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

      const entry = { nome: aluno.nome, turma: chave, media, freqMedia, sit, nivelFreq };

      turmaAlunos.push(entry);

      // Tier
      if (freqMedia !== null && freqMedia >= 90 && media !== null) {
        if (media >= 9) tierOuro.push(entry);
        else if (media >= 8) tierPrata.push(entry);
        else if (media >= 7) tierBronze.push(entry);
      }

      // Frequência perfeita
      if (freqMedia !== null && freqMedia >= 100) freqPerfeita.push(entry);

      // Alertas (modo conselho)
      if (sit === 'Reprovado' || nivelFreq === 'critico' || nivelFreq === 'alerta') {
        alertas.push({ ...entry, motivo: sit === 'Reprovado' ? 'Reprovado por nota' : nivelFreq === 'critico' ? 'Frequência crítica' : 'Frequência em alerta' });
      }
    }

    // Top destaques por turma (melhor média com freq ≥ 90)
    const turmaDestaques = turmaAlunos
      .filter((a) => a.freqMedia !== null && a.freqMedia >= 90 && a.media !== null)
      .sort((a, b) => b.media - a.media)
      .slice(0, 5);
    destaques.push(...turmaDestaques.map((d) => ({ ...d, turmaChave: chave })));
  }

  // Coletar dados por área
  for (const chave of chaves) {
    const turma = getTurma(chave);
    for (const aluno of turma.alunos) {
      for (const disc of turma.disciplinas) {
        const area = config.areas[disc.nome] || 'NÃO MAPEADA';
        if (!dadosArea[area]) {
          dadosArea[area] = {
            disciplinas: new Set(),
            totalNotas: 0, countNotas: 0,
            aprovados: 0, reprovados: 0,
            destaques: [],
            alertas: [],
          };
        }
        dadosArea[area].disciplinas.add(disc.nome);

        const dados = aluno.disciplinas[disc.nome];
        if (!dados) continue;
        const nota = resolverNota(dados.nota, config);
        if (nota !== null) {
          dadosArea[area].totalNotas += nota;
          dadosArea[area].countNotas++;
          if (nota >= config.params.mediaAprov) dadosArea[area].aprovados++;
          else dadosArea[area].reprovados++;
        }
      }
    }
  }

  // Ordenar tiers
  tierOuro.sort((a, b) => b.media - a.media);
  tierPrata.sort((a, b) => b.media - a.media);
  tierBronze.sort((a, b) => b.media - a.media);

  // Ordenar alertas por criticidade
  const critOrder = { 'Reprovado por nota': 0, 'Frequência crítica': 1, 'Frequência em alerta': 2 };
  alertas.sort((a, b) => (critOrder[a.motivo] || 3) - (critOrder[b.motivo] || 3) || a.nome.localeCompare(b.nome));

  freqPerfeita.sort((a, b) => a.nome.localeCompare(b.nome));

  return { destaques, tierOuro, tierPrata, tierBronze, freqPerfeita, alertas, dadosArea };
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function criarCapa(pptx, config, chaves) {
  const slide = pptx.addSlide();
  slide.background = { fill: '1B2838' };

  slide.addText(config.escola || 'Conselho de Classe', {
    x: 1, y: 1.5, w: '90%', h: 1,
    fontSize: 24, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Georgia',
  });
  slide.addText('Conselho de Classe — 1º Bimestre', {
    x: 1, y: 2.7, w: '90%', h: 0.6,
    fontSize: 18, color: 'C26547', align: 'center',
  });
  slide.addText(`${chaves.length} turma(s) · ${new Date().toLocaleDateString('pt-BR')}`, {
    x: 1, y: 3.5, w: '90%', h: 0.5,
    fontSize: 12, color: '8B95A3', align: 'center',
  });
}

function criarSlideDestaquesTurma(pptx, chave, config) {
  const turma = getTurma(chave);
  const alunos = turma.alunos.map((aluno) => {
    const media = calcMedia(aluno, turma.disciplinas, config);
    const freqMedia = calcFreqMedia(aluno, turma.disciplinas);
    return { nome: aluno.nome, media, freqMedia };
  });

  // Top 5 com freq ≥ 90
  const destaques = alunos
    .filter((a) => a.freqMedia !== null && a.freqMedia >= 90 && a.media !== null)
    .sort((a, b) => b.media - a.media)
    .slice(0, 6);

  if (destaques.length === 0) return;

  const slide = pptx.addSlide();
  slide.addText(`Alunos Destaques — Turma ${chave}`, {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 22, bold: true, color: '1B2838', fontFace: 'Georgia',
  });

  const rows = [[
    { text: 'Aluno', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10 } },
    { text: 'Média', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10, align: 'center' } },
    { text: 'Frequência', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10, align: 'center' } },
  ]];

  for (const d of destaques) {
    const freqStr = d.freqMedia !== null ? d.freqMedia.toFixed(0) + '%' : '—';
    const freqColor = d.freqMedia !== null && d.freqMedia >= 100 ? '4D8C62' : '131C2C';
    rows.push([
      { text: d.nome, options: { fontSize: 11, bold: true } },
      { text: d.media !== null ? d.media.toFixed(1) : '—', options: { fontSize: 12, align: 'center', bold: true, color: 'C26547' } },
      { text: freqStr, options: { fontSize: 10, align: 'center', color: freqColor } },
    ]);
  }

  const total = turma.alunos.length;
  let aprov = 0;
  for (const aluno of turma.alunos) {
    const { sit } = calcSituacao(aluno, turma, config);
    if (sit === 'Aprovado') aprov++;
  }

  slide.addText(`${total} alunos · ${aprov} aprovados (${total > 0 ? Math.round((aprov / total) * 100) : 0}%)`, {
    x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 10, color: '5A6575',
  });

  slide.addTable(rows, {
    x: 1, y: 1.7, w: 8,
    border: { type: 'solid', pt: 0.5, color: 'E0DCD5' },
    rowH: 0.45,
  });
}

function criarSlideFreqPerfeita(pptx, lista) {
  const slide = pptx.addSlide();
  slide.addText('Alunos com Destaque de Frequência', {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 22, bold: true, color: '1B2838', fontFace: 'Georgia',
  });
  slide.addText('Frequência 100% em todas as disciplinas', {
    x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 12, color: '5A6575',
  });

  const rows = [[
    { text: 'Aluno', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10 } },
    { text: 'Turma', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10, align: 'center' } },
  ]];

  for (const d of lista) {
    rows.push([
      { text: d.nome, options: { fontSize: 10 } },
      { text: d.turma, options: { fontSize: 10, align: 'center' } },
    ]);
  }

  slide.addTable(rows, {
    x: 1.5, y: 1.6, w: 7,
    border: { type: 'solid', pt: 0.5, color: 'E0DCD5' },
    rowH: 0.38,
  });
}

function criarSlideTier(pptx, titulo, lista, descricao, cor) {
  const slide = pptx.addSlide();
  slide.addText(titulo, {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 24, bold: true, color: cor, fontFace: 'Georgia',
  });
  slide.addText(descricao, {
    x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 12, color: '5A6575',
  });

  const rows = [[
    { text: 'Aluno', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10 } },
    { text: 'Turma', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10, align: 'center' } },
    { text: 'Média', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10, align: 'center' } },
    { text: 'Frequência', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 10, align: 'center' } },
  ]];

  for (const d of lista) {
    rows.push([
      { text: d.nome, options: { fontSize: 10, bold: true } },
      { text: d.turma, options: { fontSize: 10, align: 'center' } },
      { text: d.media !== null ? d.media.toFixed(1) : '—', options: { fontSize: 10, align: 'center' } },
      { text: d.freqMedia !== null ? d.freqMedia.toFixed(0) + '%' : '—', options: { fontSize: 10, align: 'center' } },
    ]);
  }

  slide.addTable(rows, {
    x: 0.5, y: 1.6, w: 9,
    border: { type: 'solid', pt: 0.5, color: 'E0DCD5' },
    rowH: 0.38,
  });
}

function criarSlideArea(pptx, areaNome, dados, config) {
  const media = dados.countNotas > 0 ? (dados.totalNotas / dados.countNotas) : 0;
  const total = dados.aprovados + dados.reprovados;
  const pctAprov = total > 0 ? Math.round((dados.aprovados / total) * 100) : 0;

  const slide = pptx.addSlide();
  slide.addText(`Área: ${areaNome}`, {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 22, bold: true, color: '1B2838', fontFace: 'Georgia',
  });

  slide.addText(
    `${dados.disciplinas.size} disciplinas · Média: ${media.toFixed(1)} · Aprovados: ${pctAprov}%`,
    { x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 12, color: '5A6575' }
  );

  // Tabela de disciplinas
  const rows = [[
    { text: 'Disciplina', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9 } },
    { text: 'Aprov.', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9, align: 'center' } },
    { text: 'Repr.', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9, align: 'center' } },
  ]];

  for (const discNome of [...dados.disciplinas].sort()) {
    let discAprov = 0, discRepr = 0;
    for (const chave of getTurmaChaves()) {
      const turma = getTurma(chave);
      const discInfo = turma.disciplinas.find((d) => d.nome === discNome);
      if (!discInfo) continue;
      for (const aluno of turma.alunos) {
        const dados = aluno.disciplinas[discNome];
        if (!dados) continue;
        const nota = resolverNota(dados.nota, config);
        if (nota !== null) {
          if (nota >= config.params.mediaAprov) discAprov++;
          else discRepr++;
        }
      }
    }
    rows.push([
      { text: discNome, options: { fontSize: 9 } },
      { text: String(discAprov), options: { fontSize: 9, align: 'center', color: '4D8C62' } },
      { text: String(discRepr), options: { fontSize: 9, align: 'center', color: discRepr > 0 ? 'C94A44' : '4D8C62' } },
    ]);
  }

  slide.addTable(rows, {
    x: 0.5, y: 1.7, w: 9,
    border: { type: 'solid', pt: 0.5, color: 'E0DCD5' },
    rowH: 0.35,
  });
}

function criarSlideAlertas(pptx, alertas, config) {
  const slide = pptx.addSlide();
  slide.addText('Encaminhamentos / Pontos de Atenção', {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 20, bold: true, color: 'C94A44', fontFace: 'Georgia',
  });

  const rows = [[
    { text: 'Aluno', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9 } },
    { text: 'Turma', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9, align: 'center' } },
    { text: 'Média', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9, align: 'center' } },
    { text: 'Freq.', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9, align: 'center' } },
    { text: 'Motivo', options: { bold: true, fill: { color: '2A3F5A' }, color: 'FFFFFF', fontSize: 9 } },
  ]];

  for (const a of alertas.slice(0, 25)) {
    rows.push([
      { text: a.nome, options: { fontSize: 9, bold: true } },
      { text: a.turma, options: { fontSize: 9, align: 'center' } },
      { text: a.media !== null ? a.media.toFixed(1) : '—', options: { fontSize: 9, align: 'center', color: 'C94A44' } },
      { text: a.freqMedia !== null ? a.freqMedia.toFixed(0) + '%' : '—', options: { fontSize: 9, align: 'center' } },
      { text: a.motivo, options: { fontSize: 9 } },
    ]);
  }

  slide.addTable(rows, {
    x: 0.3, y: 1.3, w: 9.4,
    border: { type: 'solid', pt: 0.5, color: 'E0DCD5' },
    rowH: 0.35,
  });
}
