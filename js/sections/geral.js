/**
 * sections/geral.js — Painel Geral consolidado de todas as turmas.
 */

import { getTurmas, getTurmaChaves, hasTurmas, getConfig } from '../store.js';
import { calcSituacao, calcMedia, calcFreqMedia } from '../calc.js';
import { resolverNota } from '../calc.js';
import { criarBarEmpilhada, criarBarHorizontal, pluginBarPct, dataset, COLORS } from '../charts.js';

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderGeral() {
  if (!hasTurmas()) {
    showGeralEmpty();
    return;
  }

  const config = getConfig();
  const chaves = getTurmaChaves();
  const turmas = getTurmas();

  // Coletar todos os alunos com suas turmas
  const dados = [];
  for (const chave of chaves) {
    const turma = turmas[chave];
    for (const aluno of turma.alunos) {
      dados.push({ aluno, turma, chave });
    }
  }

  // -----------------------------------------------------------------------
  // Indicadores (cards)
  // -----------------------------------------------------------------------
  let totalAprovados = 0;
  let totalReprovados = 0;
  let alertaFreq = 0;
  let critFreq = 0;
  const situacaoPorTurma = {};   // { chave: { aprovado: N, reprovado: N } }

  for (const { aluno, turma, chave } of dados) {
    const { sit, nivelFreq } = calcSituacao(aluno, turma, config);

    if (!situacaoPorTurma[chave]) situacaoPorTurma[chave] = { aprovado: 0, reprovado: 0 };
    if (sit === 'Aprovado') {
      totalAprovados++;
      situacaoPorTurma[chave].aprovado++;
    } else {
      totalReprovados++;
      situacaoPorTurma[chave].reprovado++;
    }

    if (nivelFreq === 'alerta') alertaFreq++;
    if (nivelFreq === 'critico') critFreq++;
  }

  const total = dados.length;
  const pctAprovados = total > 0 ? Math.round((totalAprovados / total) * 100) : 0;

  document.getElementById('geral-total-alunos').textContent = total;
  document.getElementById('geral-aprovados').textContent = totalAprovados;
  document.getElementById('geral-pct-aprovados').textContent = `${pctAprovados}% de ${total}`;
  document.getElementById('geral-alerta-freq').textContent = alertaFreq;
  document.getElementById('geral-crit-freq').textContent = critFreq;

  // -----------------------------------------------------------------------
  // Gráfico 1: Situação por Turma (stacked bar)
  // -----------------------------------------------------------------------
  const labels1 = chaves;
  criarBarEmpilhada('chart-geral-situacao-turma', {
    labels: labels1,
    datasets: [
      dataset('Aprovado', labels1.map((c) => situacaoPorTurma[c]?.aprovado || 0), COLORS.success, { borderRadius: 6 }),
      dataset('Reprovado', labels1.map((c) => situacaoPorTurma[c]?.reprovado || 0), COLORS.critical, { borderRadius: 6 }),
    ],
    plugins: [{ id: 'stackLabels', afterDatasetsDraw(chart) {
      const { ctx } = chart;
      chart.data.datasets.forEach((ds, i) => {
        const meta = chart.getDatasetMeta(i);
        if (meta.hidden) return;
        meta.data.forEach((bar, idx) => {
          const v = ds.data[idx];
          if (v > 0) {
            ctx.fillStyle = '#FFF';
            ctx.font = "600 11px 'Source Sans 3', sans-serif";
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(v, bar.x, bar.y + bar.height / 2);
          }
        });
      });
    }}],
  });

  // -----------------------------------------------------------------------
  // Gráfico 2: Média por Disciplina (horizontal bar)
  // -----------------------------------------------------------------------
  const todasDiscs = new Set();
  for (const turma of Object.values(turmas)) {
    for (const disc of turma.disciplinas) todasDiscs.add(disc.nome);
  }

  const discLabels = [];
  const discMedias = [];
  const discColors = [];

  for (const discNome of [...todasDiscs].sort()) {
    let soma = 0;
    let count = 0;
    for (const turma of Object.values(turmas)) {
      const discInfo = turma.disciplinas.find((d) => d.nome === discNome);
      if (!discInfo) continue;
      for (const aluno of turma.alunos) {
        const dadosDisc = aluno.disciplinas[discNome];
        if (!dadosDisc) continue;
        const nota = resolverNota(dadosDisc.nota, config);
        if (nota !== null) { soma += nota; count++; }
      }
    }
    if (count > 0) {
      discLabels.push(discNome);
      discMedias.push(soma / count);
      discColors.push((soma / count) >= config.params.mediaAprov ? COLORS.success : COLORS.critical);
    }
  }

  criarBarHorizontal('chart-geral-media-disciplina', {
    labels: discLabels,
    datasets: [
      { label: 'Média', data: discMedias, backgroundColor: discColors, borderRadius: 4, borderSkipped: false },
    ],
    ymax: 10,
    suffix: '',
    plugins: [pluginBarPct],
  });

  // -----------------------------------------------------------------------
  // Gráfico 3: Frequência Média por Turma (bar)
  // -----------------------------------------------------------------------
  const freqLabels = chaves;
  const freqMedias = [];
  const freqColors = [];

  for (const chave of chaves) {
    const turma = turmas[chave];
    let totalFreq = 0;
    let totalCount = 0;
    for (const aluno of turma.alunos) {
      const fm = calcFreqMedia(aluno, turma.disciplinas);
      if (fm !== null) { totalFreq += fm; totalCount++; }
    }
    const media = totalCount > 0 ? totalFreq / totalCount : 0;
    freqMedias.push(media);
    freqColors.push(media >= 75 ? COLORS.success : media >= 50 ? COLORS.warning : COLORS.critical);
  }

  criarBarHorizontal('chart-geral-freq-turma', {
    labels: freqLabels,
    datasets: [
      { label: 'Frequência', data: freqMedias, backgroundColor: freqColors, borderRadius: 4, borderSkipped: false },
    ],
    ymax: 100,
    suffix: '%',
    plugins: [pluginBarPct],
  });

  // -----------------------------------------------------------------------
  // Gráfico 4: Situação por Área (stacked bar)
  // -----------------------------------------------------------------------
  const areasMap = {};
  for (const chave of chaves) {
    const turma = turmas[chave];
    for (const disc of turma.disciplinas) {
      const area = config.areas[disc.nome] || 'NÃO MAPEADA';
      if (!areasMap[area]) areasMap[area] = { aprovado: 0, reprovado: 0, total: 0 };
    }
  }

  // Atribuir cada aluno-disciplina a uma área
  for (const chave of chaves) {
    const turma = turmas[chave];
    for (const disc of turma.disciplinas) {
      const area = config.areas[disc.nome] || 'NÃO MAPEADA';
      for (const aluno of turma.alunos) {
        const dadosDisc = aluno.disciplinas[disc.nome];
        if (!dadosDisc) continue;
        const nota = resolverNota(dadosDisc.nota, config);
        if (nota !== null) {
          areasMap[area].total++;
          if (nota >= config.params.mediaAprov) areasMap[area].aprovado++;
          else areasMap[area].reprovado++;
        }
      }
    }
  }

  const areaLabels = Object.keys(areasMap).sort();
  criarBarEmpilhada('chart-geral-area', {
    labels: areaLabels,
    datasets: [
      dataset('Aprovado', areaLabels.map((a) => areasMap[a].aprovado), COLORS.success, { borderRadius: 4 }),
      dataset('Reprovado', areaLabels.map((a) => areasMap[a].reprovado), COLORS.critical, { borderRadius: 4 }),
    ],
  });
}

// ---------------------------------------------------------------------------
// Empty state
// ---------------------------------------------------------------------------

function showGeralEmpty() {
  document.getElementById('geral-total-alunos').textContent = '—';
  document.getElementById('geral-aprovados').textContent = '—';
  document.getElementById('geral-pct-aprovados').textContent = '';
  document.getElementById('geral-alerta-freq').textContent = '—';
  document.getElementById('geral-crit-freq').textContent = '—';
}
