/**
 * pdf.js — Geração de relatórios PDF via jsPDF + jsPDF-AutoTable.
 *
 * jsPDF e jsPDF-AutoTable devem estar disponíveis como globais (CDN).
 */

import { getTurma, getTurmaChaves, hasTurmas, getConfig } from './store.js';
import { calcSituacao, calcMedia, calcFreqMedia, calcFreq, resolverNota } from './calc.js';

const _jsPDF = typeof jspdf !== 'undefined' ? jspdf : globalThis.jspdf;

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

export function initPDF() {
  const btn = document.getElementById('btn-gerar-pdf');

  if (btn) {
    btn.addEventListener('click', () => {
      const tipo = document.getElementById('pdf-tipo')?.value || 'geral';
      const turmaFiltro = document.getElementById('pdf-turma')?.value || 'todas';
      gerarPDF(tipo, turmaFiltro);
    });
  }

  window.addEventListener('cc:data-changed', () => {
    populatePDFTurmaSelector();
  });
}

// ---------------------------------------------------------------------------
// Gerar PDF
// ---------------------------------------------------------------------------

function gerarPDF(tipo, filtroTurma = 'todas') {
  if (!hasTurmas()) {
    alert('Importe os mapões antes de gerar relatórios.');
    return;
  }

  if (tipo === 'radar') {
    gerarPDFRadar(filtroTurma);
  } else if (tipo === 'turma') {
    gerarPDFTurma();
  } else {
    gerarPDFGeral(filtroTurma);
  }
}

// ---------------------------------------------------------------------------
// PDF Geral
// ---------------------------------------------------------------------------

function gerarPDFGeral(filtroTurma = 'todas') {
  const { jsPDF } = _jsPDF;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm' });
  const config = getConfig();
  const todasChaves = getTurmaChaves();
  const chaves = filtroTurma === 'todas' ? todasChaves : [filtroTurma];

  // Capa
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text(config.escola || 'Conselho de Classe', 15, 20);

  doc.setFontSize(14);
  doc.setFont('helvetica', 'normal');
  doc.text('Relatório Geral — Conselho de Classe', 15, 30);

  doc.setFontSize(10);
  doc.text(`Gerado em: ${new Date().toLocaleDateString('pt-BR')}`, 15, 38);
  doc.text(`Turmas: ${chaves.join(', ')}`, 15, 44);

  // Resumo consolidado
  let totalAlunos = 0, totalAprov = 0, totalRepr = 0, alertaFreq = 0, critFreq = 0;
  for (const chave of chaves) {
    const turma = getTurma(chave);
    for (const aluno of turma.alunos) {
      totalAlunos++;
      const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
      if (sit === 'Aprovado') totalAprov++;
      else totalRepr++;
      if (nivelFreq === 'alerta') alertaFreq++;
      if (nivelFreq === 'critico') critFreq++;
    }
  }

  doc.setFontSize(11);
  doc.setFont('helvetica', 'bold');
  doc.text('Indicadores Consolidados', 15, 56);

  const indicadores = [
    ['Total de Alunos Ativos', String(totalAlunos)],
    ['Aprovados', `${totalAprov} (${totalAlunos > 0 ? Math.round((totalAprov / totalAlunos) * 100) : 0}%)`],
    ['Reprovados', `${totalRepr} (${totalAlunos > 0 ? Math.round((totalRepr / totalAlunos) * 100) : 0}%)`],
    ['Alerta de Frequência (< 75%)', String(alertaFreq)],
    ['Frequência Crítica (< 50%)', String(critFreq)],
  ];

  doc.autoTable({
    startY: 60,
    head: [['Indicador', 'Valor']],
    body: indicadores,
    theme: 'grid',
    headStyles: { fillColor: [42, 63, 90], textColor: 255, fontStyle: 'bold' },
    styles: { fontSize: 9, cellPadding: 3 },
    columnStyles: { 1: { halign: 'center' } },
  });

  // Tabela por turma
  let y = doc.lastAutoTable.finalY + 10;

  for (const chave of chaves) {
    const turma = getTurma(chave);

    if (y > 170) {
      doc.addPage();
      y = 20;
    }

    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`Turma ${chave} — ${turma.alunos.length} alunos`, 15, y);
    y += 6;

    const rows = turma.alunos.map((aluno) => {
      const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
      const media = calcMedia(aluno, turma.disciplinas, config);
      const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

      let sitLabel = sit;
      if (sit === 'Aprovado' && nivelFreq === 'alerta') sitLabel = 'Aprovado ⚠ freq.';
      else if (sit === 'Aprovado' && nivelFreq === 'critico') sitLabel = 'Aprovado ⚠ retenção';

      return [
        aluno.nome,
        sitLabel,
        media !== null ? media.toFixed(1) : '—',
        freqMedia !== null ? freqMedia.toFixed(1) + '%' : '—',
      ];
    });

    doc.autoTable({
      startY: y,
      head: [['Nome', 'Situação', 'Média', 'Frequência']],
      body: rows,
      theme: 'striped',
      headStyles: { fillColor: [194, 101, 71], textColor: 255, fontStyle: 'bold' },
      styles: { fontSize: 8, cellPadding: 2 },
      columnStyles: {
        0: { cellWidth: 90 },
        1: { halign: 'center', cellWidth: 45 },
        2: { halign: 'center', cellWidth: 25 },
        3: { halign: 'center', cellWidth: 28 },
      },
    });

    y = doc.lastAutoTable.finalY + 8;
  }

  doc.save('relatorio-geral-conselho.pdf');
}

// ---------------------------------------------------------------------------
// PDF por Turma
// ---------------------------------------------------------------------------

function gerarPDFTurma() {
  const select = document.getElementById('pdf-turma');
  const chave = select?.value || getTurmaChaves()[0];
  const turma = getTurma(chave);
  const config = getConfig();

  if (!turma) return;

  const { jsPDF } = _jsPDF;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm' });
  const alunos = turma.alunos;

  // Cabeçalho
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.text(`Turma ${chave} — Relatório Detalhado`, 15, 20);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.text(`Escola: ${config.escola || 'Conselho de Classe'}`, 15, 28);
  doc.text(`Gerado em: ${new Date().toLocaleDateString('pt-BR')}`, 15, 34);

  // Indicadores
  let aprov = 0, alertaF = 0, somaMedia = 0, countMedia = 0;
  for (const aluno of alunos) {
    const { sit, nivelFreq } = calcSituacao(aluno, turma, config);
    if (sit === 'Aprovado') aprov++;
    if (nivelFreq === 'alerta') alertaF++;
    const m = calcMedia(aluno, turma.disciplinas, config);
    if (m !== null) { somaMedia += m; countMedia++; }
  }

  doc.setFontSize(10);
  doc.text(`Alunos: ${alunos.length} | Aprovados: ${aprov} | Média da Turma: ${countMedia > 0 ? (somaMedia / countMedia).toFixed(1) : '—'} | Alerta Freq.: ${alertaF}`, 15, 42);

  // Tabela com todas as disciplinas
  const headRow = ['Nome', 'Média', 'Freq.', ...turma.disciplinas.map((d) => d.nome.substring(0, 6))];

  const bodyRows = alunos.map((aluno) => {
    const media = calcMedia(aluno, turma.disciplinas, config);
    const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

    const discCells = turma.disciplinas.map((d) => {
      const dados = aluno.disciplinas[d.nome];
      if (!dados) return '—';
      const nota = resolverNota(dados.nota, config);
      return nota !== null ? (typeof nota === 'number' ? nota.toFixed(1) : nota) : '—';
    });

    return [
      aluno.nome,
      media !== null ? media.toFixed(1) : '—',
      freqMedia !== null ? freqMedia.toFixed(1) + '%' : '—',
      ...discCells,
    ];
  });

  doc.autoTable({
    startY: 48,
    head: [headRow],
    body: bodyRows,
    theme: 'grid',
    headStyles: { fillColor: [42, 63, 90], textColor: 255, fontStyle: 'bold' },
    styles: { fontSize: 7, cellPadding: 2 },
    columnStyles: {
      0: { cellWidth: 70 },
      1: { halign: 'center', cellWidth: 18 },
      2: { halign: 'center', cellWidth: 20 },
    },
  });

  doc.save(`relatorio-turma-${chave}.pdf`);
}

// ---------------------------------------------------------------------------
// PDF Radar
// ---------------------------------------------------------------------------

function gerarPDFRadar(filtroTurma = 'todas') {
  const config = getConfig();
  const todasChaves = getTurmaChaves();
  const chaves = filtroTurma === 'todas' ? todasChaves : [filtroTurma];

  // Coletar alunos em risco (mesma lógica do radar.js)
  const riscos = [];
  for (const chave of chaves) {
    const turma = getTurma(chave);
    for (const aluno of turma.alunos) {
      const { sit, nivelFreq } = calcSituacao(aluno, turma, config);

      let reps = 0;
      for (const disc of turma.disciplinas) {
        const dados = aluno.disciplinas[disc.nome];
        if (!dados) continue;
        const nota = resolverNota(dados.nota, config);
        if (nota !== null && nota < config.params.mediaAprov) reps++;
      }

      let criticidade = 0;
      const motivos = [];
      if (nivelFreq === 'critico') { criticidade = 3; motivos.push('Frequência crítica (< 50%)'); }
      if (reps >= 2) { criticidade = Math.max(criticidade, 3); motivos.push(`Reprovado em ${reps} disciplinas`); }
      if (nivelFreq === 'alerta' && criticidade < 2) { criticidade = Math.max(criticidade, 2); motivos.push('Frequência em alerta (< 75%)'); }
      if (sit === 'Reprovado' && criticidade < 2) { criticidade = Math.max(criticidade, 2); motivos.push('Reprovado por nota'); }

      if (criticidade > 0) {
        riscos.push({ nome: aluno.nome, turma: chave, criticidade, motivos: motivos.join(' · ') });
      }
    }
  }
  riscos.sort((a, b) => b.criticidade - a.criticidade || a.nome.localeCompare(b.nome));

  const { jsPDF } = _jsPDF;
  const doc = new jsPDF();
  const config2 = getConfig();

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.text('Radar Preditivo — Alunos em Risco', 15, 20);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.text(`Escola: ${config2.escola || 'Conselho de Classe'}`, 15, 28);
  doc.text(`Gerado em: ${new Date().toLocaleDateString('pt-BR')}`, 15, 34);

  if (riscos.length === 0) {
    doc.text('Nenhum aluno em risco identificado.', 15, 44);
  } else {
    const rows = riscos.map((r) => [
      r.nome,
      `Turma ${r.turma}`,
      r.criticidade >= 3 ? 'Crítico' : 'Alerta',
      r.motivos,
    ]);

    doc.autoTable({
      startY: 40,
      head: [['Nome', 'Turma', 'Nível', 'Motivos']],
      body: rows,
      theme: 'grid',
      headStyles: { fillColor: [194, 101, 71], textColor: 255, fontStyle: 'bold' },
      styles: { fontSize: 9, cellPadding: 3 },
      columnStyles: {
        0: { cellWidth: 60 },
        1: { halign: 'center', cellWidth: 25 },
        2: { halign: 'center', cellWidth: 25 },
        3: { cellWidth: 'auto' },
      },
    });
  }

  doc.save('radar-preditivo.pdf');
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function populatePDFTurmaSelector() {
  const select = document.getElementById('pdf-turma');
  if (!select) return;
  const current = select.value || 'todas';
  const chaves = getTurmaChaves();
  select.innerHTML = '<option value="todas">Todas as turmas</option>' +
    chaves.map((c) => `<option value="${c}" ${c === current ? 'selected' : ''}>Turma ${c}</option>`).join('');
}
