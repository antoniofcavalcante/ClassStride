/**
 * sections/radar.js — Radar Preditivo de alunos em risco.
 */

import { getTurmaChaves, getTurma, getTurmaAnterior, hasTurmas, hasTurmasAnterior, getConfig } from '../store.js';
import { calcSituacao, calcFreqMedia, calcMedia, resolverNota } from '../calc.js';

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

export function initRadar() {
  const select = document.getElementById('select-radar-turma');
  if (!select) return;

  select.addEventListener('change', () => renderRadar());

  window.addEventListener('cc:data-changed', () => {
    populateRadarSelector(select);
    renderRadar();
  });
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

export function renderRadar() {
  if (!hasTurmas()) {
    showEmpty();
    return;
  }

  const config = getConfig();
  const select = document.getElementById('select-radar-turma');
  populateRadarSelector(select);

  const filtroTurma = select?.value || 'todas';
  const chaves = filtroTurma === 'todas' ? getTurmaChaves() : [filtroTurma];

  // Coletar TODOS os alunos e avaliar risco
  const riscos = [];

  for (const chave of chaves) {
    const turma = getTurma(chave);
    if (!turma) continue;

    for (const aluno of turma.alunos) {
      const { sit, nivelFreq } = calcSituacao(aluno, turma, config);

      // Contar reprovações
      let reps = 0;
      for (const disc of turma.disciplinas) {
        const dados = aluno.disciplinas[disc.nome];
        if (!dados) continue;
        const nota = resolverNota(dados.nota, config);
        if (nota !== null && nota < config.params.mediaAprov) reps++;
      }

      // Determinar nível de risco
      let criticidade = 0; // 0=nenhum, 1=atenção, 2=alerta, 3=crítico
      const motivos = [];

      if (nivelFreq === 'critico') {
        criticidade = 3;
        motivos.push('Frequência crítica (< 50%)');
      }

      if (reps >= 2) {
        criticidade = Math.max(criticidade, 3);
        motivos.push(`Reprovado em ${reps} disciplinas`);
      }

      if (nivelFreq === 'alerta' && criticidade < 2) {
        criticidade = Math.max(criticidade, 2);
        motivos.push('Frequência em alerta (< 75%)');
      }

      if (sit === 'Reprovado' && criticidade < 2) {
        criticidade = Math.max(criticidade, 2);
        if (!motivos.some((m) => m.includes('Reprovado'))) {
          motivos.push('Reprovado por nota');
        }
      }

      if (criticidade > 0) {
        const freqMedia = calcFreqMedia(aluno, turma.disciplinas);
        const media = calcMedia(aluno, turma.disciplinas, config);

        riscos.push({
          nome: aluno.nome,
          turma: chave,
          criticidade,
          motivos,
          media,
          freqMedia,
          sit,
          nivelFreq,
        });
      }
    }
  }

  // Ordenar por criticidade (desc), depois por nome
  riscos.sort((a, b) => b.criticidade - a.criticidade || a.nome.localeCompare(b.nome));

  // Renderizar
  const listEl = document.getElementById('radar-list');
  const emptyEl = document.getElementById('radar-empty');

  if (riscos.length === 0) {
    if (listEl) listEl.innerHTML = '';
    if (emptyEl) emptyEl.classList.remove('hidden');
    return;
  }

  if (emptyEl) emptyEl.classList.add('hidden');

  if (listEl) {
    listEl.innerHTML = riscos
      .map(
        (r, i) => {
          const rowClass = r.criticidade >= 3 ? 'radar-item-critical' : 'radar-item-warning';
          const criticidadeLabel = r.criticidade >= 3 ? 'Crítico' : 'Alerta';

          const badgesHtml = [
            r.sit === 'Reprovado' ? '<span class="badge badge-critical">Reprovado</span>' : '',
            r.nivelFreq === 'critico'
              ? '<span class="badge badge-critical">Freq. Crítica</span>'
              : r.nivelFreq === 'alerta'
                ? '<span class="badge badge-warning">Freq. Alerta</span>'
                : '',
          ]
            .filter(Boolean)
            .join('');

          // Evolução vs bimestre anterior
          let evolucaoHtml = '';
          if (hasTurmasAnterior()) {
            const turmaAnt = getTurmaAnterior(r.turma);
            if (turmaAnt) {
              const alunoAnt = turmaAnt.alunos.find((a) => a.nome === r.nome);
              if (alunoAnt) {
                // Média anterior
                let mediaAnt = null, freqAnt = null;
                let soma = 0, count = 0;
                for (const disc of turmaAnt.disciplinas) {
                  const dados = alunoAnt.disciplinas[disc.nome];
                  if (!dados) continue;
                  if (typeof dados.nota === 'number') { soma += dados.nota; count++; }
                }
                mediaAnt = count > 0 ? soma / count : null;

                let totalAulas = 0, totalPresente = 0;
                for (const disc of turmaAnt.disciplinas) {
                  const dados = alunoAnt.disciplinas[disc.nome];
                  if (!dados || disc.aulasDadas === 0) continue;
                  totalAulas += disc.aulasDadas;
                  totalPresente += (disc.aulasDadas - dados.faltas);
                }
                freqAnt = totalAulas > 0 ? (totalPresente / totalAulas) * 100 : null;

                if (mediaAnt !== null && r.media !== null) {
                  const diffMedia = r.media - mediaAnt;
                  if (Math.abs(diffMedia) >= 0.1) {
                    const up = diffMedia > 0;
                    evolucaoHtml += `<span style="color:${up ? '#4D8C62' : '#C94A44'}; font-size:0.85em; margin-left:4px;">${up ? '↑' : '↓'}${Math.abs(diffMedia).toFixed(1)}</span>`;
                  }
                }
                if (freqAnt !== null && r.freqMedia !== null) {
                  const diffFreq = r.freqMedia - freqAnt;
                  if (Math.abs(diffFreq) >= 1) {
                    const up = diffFreq > 0;
                    evolucaoHtml += `<span style="color:${up ? '#4D8C62' : '#C94A44'}; font-size:0.85em; margin-left:4px;">${up ? '↑' : '↓'}${Math.abs(diffFreq).toFixed(0)}%</span>`;
                  }
                }
              }
            }
          }

          return `
          <div class="radar-item ${rowClass}">
            <div class="radar-item-priority">${i + 1}</div>
            <div class="radar-item-info">
              <div class="radar-item-nome">${esc(r.nome)}${evolucaoHtml}</div>
              <div class="radar-item-turma">Turma ${r.turma} · Média: ${r.media !== null ? r.media.toFixed(1) : '—'} · Freq.: ${r.freqMedia !== null ? r.freqMedia.toFixed(1) + '%' : '—'}</div>
              <div class="radar-item-motivo">${r.motivos.join(' · ')}</div>
            </div>
            <div class="radar-item-badges">${badgesHtml}</div>
          </div>`;
        }
      )
      .join('');
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function populateRadarSelector(select) {
  if (!select) return;
  const current = select.value || 'todas';
  const chaves = getTurmaChaves();
  select.innerHTML =
    '<option value="todas">Todas as turmas</option>' +
    chaves.map((c) => `<option value="${c}" ${c === current ? 'selected' : ''}>Turma ${c}</option>`).join('');
}

function showEmpty() {
  const listEl = document.getElementById('radar-list');
  const emptyEl = document.getElementById('radar-empty');
  if (listEl) listEl.innerHTML = '';
  if (emptyEl) emptyEl.classList.remove('hidden');
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
