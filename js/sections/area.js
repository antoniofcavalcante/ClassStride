/**
 * sections/area.js — Análise por área configurada.
 *
 * Visão detalhada por área com destaque de alunos (ouro/prata/bronze),
 * tabela de disciplinas e alertas — ideal para projeção em reunião.
 */

import { getTurmaChaves, getTurma, hasTurmas, getConfig } from '../store.js';
import { resolverNota, calcFreq, calcMedia, calcFreqMedia } from '../calc.js';
import { criarBarHorizontal, criarBarChart, pluginBarPct, COLORS } from '../charts.js';

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderArea() {
  const contentEl = document.getElementById('area-content');
  const emptyEl = document.getElementById('area-empty');

  if (!hasTurmas()) {
    if (contentEl) contentEl.innerHTML = '';
    if (emptyEl) emptyEl.classList.remove('hidden');
    return;
  }

  const config = getConfig();
  const chaves = getTurmaChaves();

  // Selector de área e turma (ler antes de usar)
  const currentArea = document.getElementById('select-area-filter')?.value || '';
  const currentTurma = document.getElementById('select-area-turma')?.value || 'todas';

  // Agrupar disciplinas por área + coletar dados de alunos
  const areas = buildAreasData(chaves, config, currentTurma);

  const areaNames = Object.keys(areas).sort();
  if (areaNames.length === 0) {
    if (contentEl) contentEl.innerHTML = '';
    if (emptyEl) emptyEl.classList.remove('hidden');
    return;
  }
  if (emptyEl) emptyEl.classList.add('hidden');

  // Área padrão se nenhuma selecionada
  const selectedArea = currentArea || areaNames[0];

  let html = '';

  // Filtros dropdown
  html += `
  <div class="card">
    <div class="card-body" style="padding: var(--space-4) var(--space-5);">
      <div class="flex items-center gap-3 flex-wrap">
        <div class="form-group" style="margin:0;">
          <label for="select-area-filter" class="form-label">Área:</label>
          <select id="select-area-filter" class="form-select" style="min-width: 260px;">
            <option value="">Todas as áreas (visão geral)</option>
            ${areaNames.map((a) => `<option value="${esc(a)}" ${a === selectedArea ? 'selected' : ''}>${esc(a)}</option>`).join('')}
          </select>
        </div>
        <div class="form-group" style="margin:0;">
          <label for="select-area-turma" class="form-label">Turma:</label>
          <select id="select-area-turma" class="form-select" style="min-width: 160px;">
            <option value="todas" ${currentTurma === 'todas' ? 'selected' : ''}>Todas as turmas</option>
            ${chaves.map((c) => `<option value="${c}" ${c === currentTurma ? 'selected' : ''}>Turma ${c}</option>`).join('')}
          </select>
        </div>
      </div>
    </div>
  </div>`;

  if (selectedArea) {
    html += renderAreaDetail(selectedArea, areas[selectedArea], config);
  } else {
    html += renderAllAreasSummary(areas, areaNames, config);
  }

  if (contentEl) {
    contentEl.innerHTML = html;

    // Re-bind selectors
    const selArea = document.getElementById('select-area-filter');
    const selTurma = document.getElementById('select-area-turma');
    if (selArea) selArea.addEventListener('change', () => renderArea());
    if (selTurma) selTurma.addEventListener('change', () => renderArea());
  }

  // Gráfico comparativo (só na visão geral)
  if (!currentArea) {
    setTimeout(() => {
      renderComparativoChart(areas, areaNames, config);
    }, 50);
  }

  // Gráfico de disciplinas na visão detalhada
  if (currentArea) {
    setTimeout(() => {
      renderAreaDiscChart(areas[selectedArea], config);
    }, 50);
  }
}

// ---------------------------------------------------------------------------
// Dados agregados por área e aluno
// ---------------------------------------------------------------------------

function buildAreasData(chaves, config, filtroTurma = 'todas') {
  const areas = {};
  const turmasFiltradas = filtroTurma === 'todas' ? chaves : [filtroTurma];

  for (const chave of turmasFiltradas) {
    const turma = getTurma(chave);

    // Para cada aluno, calcular métricas gerais
    for (const aluno of turma.alunos) {
      const media = calcMedia(aluno, turma.disciplinas, config);
      const freqMedia = calcFreqMedia(aluno, turma.disciplinas);

      // Classificar tier: ouro/prata/bronze
      let tier = null;
      if (freqMedia !== null && freqMedia >= 90 && media !== null) {
        if (media >= 9) tier = 'ouro';
        else if (media >= 8) tier = 'prata';
        else if (media >= 7) tier = 'bronze';
      }

      // Para cada disciplina → área
      for (const disc of turma.disciplinas) {
        const area = config.areas[disc.nome] || 'NÃO MAPEADA';
        if (!areas[area]) {
          areas[area] = {
            disciplinas: new Set(),
            totalNotas: 0, countNotas: 0,
            totalFreq: 0, countFreq: 0,
            discStats: {},
            // Destaques por área (baseado nas disciplinas dessa área)
            destaques: [],
            alertas: [],
            alunos: new Map(), // alunoKey → { nome, turma, notas[], freqs[] }
          };
        }

        areas[area].disciplinas.add(disc.nome);

        // Init disc stats
        if (!areas[area].discStats[disc.nome]) {
          areas[area].discStats[disc.nome] = {
            somaNota: 0, countNota: 0, aprov: 0, repr: 0,
            somaFreq: 0, countFreq: 0, aulasDadas: disc.aulasDadas,
          };
        }

        const dados = aluno.disciplinas[disc.nome];
        if (!dados) continue;

        const nota = resolverNota(dados.nota, config);
        const freq = calcFreq(dados.faltas, disc.aulasDadas);

        if (nota !== null) {
          areas[area].totalNotas += nota;
          areas[area].countNotas++;
          areas[area].discStats[disc.nome].somaNota += nota;
          areas[area].discStats[disc.nome].countNota++;

          if (nota >= config.params.mediaAprov) {
            areas[area].discStats[disc.nome].aprov++;
          } else {
            areas[area].discStats[disc.nome].repr++;
          }
        }

        if (freq !== null) {
          areas[area].totalFreq += freq;
          areas[area].countFreq++;
          areas[area].discStats[disc.nome].somaFreq += freq;
          areas[area].discStats[disc.nome].countFreq++;
        }

        // Acumular por aluno na área
        const key = `${chave}|${aluno.nome}`;
        if (!areas[area].alunos.has(key)) {
          areas[area].alunos.set(key, { nome: aluno.nome, turma: chave, notas: [], freqs: [] });
        }
        const aa = areas[area].alunos.get(key);
        if (nota !== null) aa.notas.push(nota);
        if (freq !== null) aa.freqs.push(freq);
      }
    }

    // Calcular destaques e alertas por área
    for (const area of Object.keys(areas)) {
      const a = areas[area];
      let aprovados = 0;
      let reprovados = 0;

      for (const [, aa] of a.alunos) {
        const mediaArea = aa.notas.length > 0 ? aa.notas.reduce((s, v) => s + v, 0) / aa.notas.length : null;
        const freqArea = aa.freqs.length > 0 ? aa.freqs.reduce((s, v) => s + v, 0) / aa.freqs.length : null;

        // Contar por ALUNO (não por avaliação)
        if (mediaArea !== null) {
          if (mediaArea >= config.params.mediaAprov) aprovados++;
          else reprovados++;
        }

        let tier = null;
        if (freqArea !== null && freqArea >= 90 && mediaArea !== null) {
          if (mediaArea >= 9) tier = 'ouro';
          else if (mediaArea >= 8) tier = 'prata';
          else if (mediaArea >= 7) tier = 'bronze';
        }

        if (mediaArea !== null && freqArea !== null) {
          a.destaques.push({ nome: aa.nome, turma: aa.turma, media: mediaArea, freqMedia: freqArea, tier });
        }

        // Alertas: nota < 5 ou freq < 75%
        if (mediaArea !== null && mediaArea < config.params.mediaAprov) {
          a.alertas.push({ nome: aa.nome, turma: aa.turma, media: mediaArea, freqMedia: freqArea, motivo: 'Nota abaixo da média' });
        } else if (freqArea !== null && freqArea < config.params.freqMin) {
          a.alertas.push({ nome: aa.nome, turma: aa.turma, media: mediaArea, freqMedia: freqArea, motivo: 'Frequência baixa' });
        }
      }

      a.aprovados = aprovados;
      a.reprovados = reprovados;
      a.totalAlunos = a.alunos.size;

      // Ordenar destaques por tier (ouro > prata > bronze) e média
      const tierOrder = { ouro: 0, prata: 1, bronze: 2 };
      a.destaques.sort((x, y) => (tierOrder[x.tier] || 3) - (tierOrder[y.tier] || 3) || y.media - x.media);

      // Ordenar alertas por criticidade
      a.alertas.sort((x, y) => x.media - y.media);
    }
  }

  return areas;
}

// ---------------------------------------------------------------------------
// Visão detalhada de UMA área
// ---------------------------------------------------------------------------

function renderAreaDetail(areaName, a, config) {
  const media = a.countNotas > 0 ? (a.totalNotas / a.countNotas) : 0;
  const freqMedia = a.countFreq > 0 ? (a.totalFreq / a.countFreq) : 0;
  const totalAlunos = a.totalAlunos || a.alunos.size;
  const pctAprov = totalAlunos > 0 ? Math.round((a.aprovados / totalAlunos) * 100) : 0;

  let html = '';

  // Cabeçalho da área
  html += `
  <div class="card">
    <div class="card-header">
      <h3>${esc(areaName)}</h3>
      <span class="text-sm text-muted">${a.disciplinas.size} disciplina(s)</span>
    </div>
    <div class="card-body">
      <div class="grid-4">
        <div class="stat-card">
          <span class="stat-card-label">Alunos na Área</span>
          <span class="stat-card-value">${totalAlunos}</span>
        </div>
        <div class="stat-card">
          <span class="stat-card-label">Média da Área</span>
          <span class="stat-card-value">${media.toFixed(1)}</span>
        </div>
        <div class="stat-card">
          <span class="stat-card-label">Frequência Média</span>
          <span class="stat-card-value">${freqMedia.toFixed(1)}%</span>
        </div>
        <div class="stat-card">
          <span class="stat-card-label">Aprovados</span>
          <span class="stat-card-value text-success">${a.aprovados} <span class="stat-card-aux">(${pctAprov}%)</span></span>
        </div>
      </div>
    </div>
  </div>`;

  // Disciplinas da área
  html += `
  <div class="card">
    <div class="card-header"><h3>Disciplinas</h3></div>
    <div class="card-body" style="padding:0;">
      <div class="table-wrapper">
        <table class="table table-compact">
          <thead>
            <tr>
              <th>Disciplina</th>
              <th class="table-numeric">Média</th>
              <th class="table-numeric">Aprov.</th>
              <th class="table-numeric">Repr.</th>
              <th class="table-numeric">Freq.</th>
              <th class="table-numeric">Aulas</th>
            </tr>
          </thead>
          <tbody>
            ${[...a.disciplinas].sort().map((discNome) => {
              const ds = a.discStats[discNome];
              const discMedia = ds.countNota > 0 ? (ds.somaNota / ds.countNota).toFixed(1) : '—';
              const discFreq = ds.countFreq > 0 ? (ds.somaFreq / ds.countFreq).toFixed(1) : '—';
              const abaixo = ds.countNota > 0 && (ds.somaNota / ds.countNota) < config.params.mediaAprov;
              return `
              <tr>
                <td>${esc(discNome)}</td>
                <td class="table-numeric ${abaixo ? 'text-critical font-semibold' : ''}">${discMedia}</td>
                <td class="table-numeric text-success">${ds.aprov}</td>
                <td class="table-numeric text-critical">${ds.repr}</td>
                <td class="table-numeric">${discFreq}%</td>
                <td class="table-numeric">${ds.aulasDadas}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>
  </div>`;

  // Gráfico comparativo de disciplinas dentro da área
  html += `
  <div class="card">
    <div class="card-header"><h3>Comparativo de Disciplinas</h3></div>
    <div class="card-body">
      <div class="chart-container chart-container-lg"><canvas id="chart-area-discs"></canvas></div>
    </div>
  </div>`;

  // Destaques Ouro / Prata / Bronze
  const ouro = a.destaques.filter((d) => d.tier === 'ouro');
  const prata = a.destaques.filter((d) => d.tier === 'prata');
  const bronze = a.destaques.filter((d) => d.tier === 'bronze');

  if (ouro.length + prata.length + bronze.length > 0) {
    html += `<div class="card"><div class="card-header"><h3>Destaques</h3></div><div class="card-body">`;

    if (ouro.length > 0) {
      html += `<h4 style="color: #C4953A; margin-bottom: var(--space-3);">🥇 Ouro (Freq ≥ 90% · Média ≥ 9)</h4>`;
      html += renderDestaqueTable(ouro);
    }
    if (prata.length > 0) {
      html += `<h4 style="color: #8B95A3; margin: var(--space-4) 0 var(--space-3);">🥈 Prata (Freq ≥ 90% · Média 8–9)</h4>`;
      html += renderDestaqueTable(prata);
    }
    if (bronze.length > 0) {
      html += `<h4 style="color: #C26547; margin: var(--space-4) 0 var(--space-3);">🥉 Bronze (Freq ≥ 90% · Média 7–8)</h4>`;
      html += renderDestaqueTable(bronze);
    }

    html += `</div></div>`;
  }

  // Alertas
  if (a.alertas.length > 0) {
    html += `
    <div class="card">
      <div class="card-header"><h3>⚠️ Pontos de Atenção</h3></div>
      <div class="card-body" style="padding:0;">
        <div class="table-wrapper">
          <table class="table table-compact">
            <thead>
              <tr>
                <th>Aluno</th>
                <th>Turma</th>
                <th class="table-numeric">Média</th>
                <th class="table-numeric">Freq.</th>
                <th>Motivo</th>
              </tr>
            </thead>
            <tbody>
              ${a.alertas.map((al) => `
              <tr>
                <td class="font-medium">${esc(al.nome)}</td>
                <td>${al.turma}</td>
                <td class="table-numeric ${al.media < config.params.mediaAprov ? 'text-critical font-semibold' : ''}">${al.media.toFixed(1)}</td>
                <td class="table-numeric ${al.freqMedia < config.params.freqMin ? 'text-warning font-semibold' : ''}">${al.freqMedia.toFixed(1)}%</td>
                <td>${al.motivo}</td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>`;
  }

  return html;
}

function renderDestaqueTable(lista) {
  return `
  <div class="table-wrapper">
    <table class="table table-compact">
      <thead>
        <tr>
          <th>Aluno</th>
          <th>Turma</th>
          <th class="table-numeric">Média</th>
          <th class="table-numeric">Frequência</th>
        </tr>
      </thead>
      <tbody>
        ${lista.map((d) => `
        <tr>
          <td class="font-semibold">${esc(d.nome)}</td>
          <td>${d.turma}</td>
          <td class="table-numeric font-semibold">${d.media.toFixed(1)}</td>
          <td class="table-numeric">${d.freqMedia.toFixed(1)}%</td>
        </tr>`).join('')}
      </tbody>
    </table>
  </div>`;
}

// ---------------------------------------------------------------------------
// Visão geral (todas as áreas)
// ---------------------------------------------------------------------------

function renderAllAreasSummary(areas, areaNames, config) {
  let html = '<div class="grid-2">';

  for (const areaName of areaNames) {
    const a = areas[areaName];
    const media = a.countNotas > 0 ? (a.totalNotas / a.countNotas) : 0;
    const freqMedia = a.countFreq > 0 ? (a.totalFreq / a.countFreq) : 0;
    const totalAlunos = a.totalAlunos || a.alunos.size;
    const pctAprov = totalAlunos > 0 ? Math.round((a.aprovados / totalAlunos) * 100) : 0;
    const isNaoMapeada = areaName === 'NÃO MAPEADA';
    const ouro = a.destaques.filter((d) => d.tier === 'ouro').length;
    const prata = a.destaques.filter((d) => d.tier === 'prata').length;
    const bronze = a.destaques.filter((d) => d.tier === 'bronze').length;

    html += `
    <div class="card">
      <div class="card-header">
        <h3>${esc(areaName)}</h3>
        ${isNaoMapeada ? '<span class="badge badge-warning">Não classificada</span>' : ''}
      </div>
      <div class="card-body">
        <div class="grid-2 mb-3">
          <div class="stat-card">
            <span class="stat-card-label">Alunos</span>
            <span class="stat-card-value">${totalAlunos}</span>
          </div>
          <div class="stat-card">
            <span class="stat-card-label">Média</span>
            <span class="stat-card-value">${media.toFixed(1)}</span>
          </div>
          <div class="stat-card">
            <span class="stat-card-label">Frequência</span>
            <span class="stat-card-value">${freqMedia.toFixed(1)}%</span>
          </div>
          <div class="stat-card">
            <span class="stat-card-label">Aprovados</span>
            <span class="stat-card-value text-success">${pctAprov}%</span>
          </div>
          <div class="stat-card">
            <span class="stat-card-label">Destaques</span>
            <span class="stat-card-value">${ouro + prata + bronze}</span>
            <span class="stat-card-aux">🥇${ouro} 🥈${prata} 🥉${bronze}</span>
          </div>
        </div>
      </div>
    </div>`;
  }

  html += '</div>';
  return html;
}

// ---------------------------------------------------------------------------
// Gráfico comparativo
// ---------------------------------------------------------------------------

function renderComparativoChart(areas, areaNames, config) {
  const labels = areaNames;
  const medias = areaNames.map((a) => {
    const s = areas[a];
    return s.countNotas > 0 ? s.totalNotas / s.countNotas : 0;
  });
  const colors = areaNames.map((a) => {
    const s = areas[a];
    const m = s.countNotas > 0 ? s.totalNotas / s.countNotas : 0;
    return m >= config.params.mediaAprov ? COLORS.success : COLORS.critical;
  });

  criarBarHorizontal('chart-area-media', {
    labels,
    datasets: [
      { label: 'Média', data: medias, backgroundColor: colors, borderRadius: 4, borderSkipped: false },
    ],
    ymax: 10,
    plugins: [pluginBarPct],
  });
}

function renderAreaDiscChart(areaData, config) {
  const discNames = [...areaData.disciplinas].sort();
  const medias = discNames.map((nome) => {
    const ds = areaData.discStats[nome];
    return ds && ds.countNota > 0 ? ds.somaNota / ds.countNota : 0;
  });
  const colors = medias.map((m) =>
    m >= config.params.mediaAprov ? COLORS.success : COLORS.critical
  );

  criarBarChart('chart-area-discs', {
    labels: discNames,
    datasets: [
      {
        label: 'Média',
        data: medias,
        backgroundColor: colors,
        borderRadius: 4,
        borderSkipped: false,
      },
    ],
    ymax: 10,
    plugins: [pluginBarPct],
  });
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
