/**
 * charts.js — Chart.js plugins e factory functions.
 *
 * Gerencia criação/destruição de instâncias Chart.js.
 * Cada chamada a um factory destrói a instância anterior do mesmo canvas.
 */

// ---------------------------------------------------------------------------
// Chart registry (para destruir instâncias anteriores)
// ---------------------------------------------------------------------------

const _charts = {};

function destroyIfExists(canvasId) {
  if (_charts[canvasId]) {
    _charts[canvasId].destroy();
    delete _charts[canvasId];
  }
}

// ---------------------------------------------------------------------------
// Plugins
// ---------------------------------------------------------------------------

/**
 * Plugin: exibe valor percentual dentro/sobre barras.
 */
export const pluginBarPct = {
  id: 'barPct',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, dsIndex) => {
      const meta = chart.getDatasetMeta(dsIndex);
      if (!meta.hidden) {
        meta.data.forEach((bar, index) => {
          const value = dataset.data[index];
          if (value == null) return;
          const label = Number.isInteger(value) ? `${value}%` : `${value.toFixed(1)}%`;

          ctx.fillStyle = '#131C2C';
          ctx.font = "600 11px 'Source Sans 3', sans-serif";
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';

          if (bar.height > 20) {
            ctx.fillText(label, bar.x, bar.y + bar.height / 2);
          } else {
            ctx.fillText(label, bar.x, bar.y - 8);
          }
        });
      }
    });
  },
};

/**
 * Plugin: exibe valor percentual sobre pontos de linha.
 */
export const pluginLinePct = {
  id: 'linePct',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, dsIndex) => {
      const meta = chart.getDatasetMeta(dsIndex);
      if (meta.hidden) return;
      meta.data.forEach((point, index) => {
        const value = dataset.data[index];
        if (value == null) return;
        const label = Number.isInteger(value) ? `${value}%` : `${value.toFixed(1)}%`;

        ctx.fillStyle = '#131C2C';
        ctx.font = "600 10px 'Source Sans 3', sans-serif";
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(label, point.x, point.y - 6);
      });
    });
  },
};

/**
 * Plugin: exibe valor numérico nas barras.
 */
export const pluginBarValor = {
  id: 'barValor',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, dsIndex) => {
      const meta = chart.getDatasetMeta(dsIndex);
      if (!meta.hidden) {
        meta.data.forEach((bar, index) => {
          const value = dataset.data[index];
          if (value == null) return;
          const label = Number.isInteger(value) ? String(value) : value.toFixed(1);

          ctx.fillStyle = '#FFFFFF';
          ctx.font = "600 11px 'Source Sans 3', sans-serif";
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';

          if (bar.height > 24) {
            ctx.fillText(label, bar.x, bar.y + 14);
          } else {
            ctx.fillText(label, bar.x, bar.y - 8);
          }
        });
      }
    });
  },
};

// ---------------------------------------------------------------------------
// Cores do tema
// ---------------------------------------------------------------------------

const COLORS = {
  primary: '#C26547',
  secondary: '#2A3F5A',
  success: '#4D8C62',
  warning: '#C4953A',
  critical: '#C94A44',
  info: '#4A607A',
};

const STATUS_COLORS = {
  aprovado: '#4D8C62',
  reprovado: '#C94A44',
  alerta: '#C4953A',
  critico: '#C94A44',
};

// ---------------------------------------------------------------------------
// Chart defaults
// ---------------------------------------------------------------------------

function baseOptions(overrides = {}) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'bottom',
        labels: {
          usePointStyle: true,
          padding: 16,
          font: { family: "'Source Sans 3', sans-serif", size: 12 },
        },
      },
    },
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Factory: Bar Chart (vertical)
// ---------------------------------------------------------------------------

export function criarBarChart(canvasId, config) {
  destroyIfExists(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;

  const ctx = canvas.getContext('2d');

  const chartConfig = {
    type: 'bar',
    data: {
      labels: config.labels || [],
      datasets: config.datasets || [],
    },
    options: baseOptions({
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          max: config.ymax,
          ticks: { callback: (v) => (config.suffix ? v + config.suffix : v) },
        },
      },
      ...config.options,
    }),
    plugins: config.plugins || [],
  };

  _charts[canvasId] = new Chart(ctx, chartConfig);
  return _charts[canvasId];
}

// ---------------------------------------------------------------------------
// Factory: Horizontal Bar Chart
// ---------------------------------------------------------------------------

export function criarBarHorizontal(canvasId, config) {
  destroyIfExists(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;

  const ctx = canvas.getContext('2d');

  const chartConfig = {
    type: 'bar',
    data: {
      labels: config.labels || [],
      datasets: config.datasets || [],
    },
    options: baseOptions({
      indexAxis: 'y',
      scales: {
        x: {
          beginAtZero: true,
          max: config.ymax,
          ticks: { callback: (v) => (config.suffix ? v + config.suffix : v) },
        },
        y: { grid: { display: false } },
      },
      ...config.options,
    }),
    plugins: config.plugins || [],
  };

  _charts[canvasId] = new Chart(ctx, chartConfig);
  return _charts[canvasId];
}

// ---------------------------------------------------------------------------
// Factory: Stacked Bar
// ---------------------------------------------------------------------------

export function criarBarEmpilhada(canvasId, config) {
  destroyIfExists(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;

  const ctx = canvas.getContext('2d');

  const chartConfig = {
    type: 'bar',
    data: {
      labels: config.labels || [],
      datasets: (config.datasets || []).map((ds) => ({ ...ds })),
    },
    options: baseOptions({
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, beginAtZero: true },
      },
      ...config.options,
    }),
    plugins: config.plugins || [],
  };

  _charts[canvasId] = new Chart(ctx, chartConfig);
  return _charts[canvasId];
}

// ---------------------------------------------------------------------------
// Factory: Doughnut / Pie
// ---------------------------------------------------------------------------

export function criarDoughnut(canvasId, config) {
  destroyIfExists(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;

  const ctx = canvas.getContext('2d');

  const chartConfig = {
    type: 'doughnut',
    data: {
      labels: config.labels || [],
      datasets: [
        {
          data: config.data || [],
          backgroundColor: config.colors || Object.values(COLORS),
          borderWidth: 2,
          borderColor: '#FFFFFF',
        },
      ],
    },
    options: baseOptions({
      cutout: '60%',
      ...config.options,
    }),
  };

  _charts[canvasId] = new Chart(ctx, chartConfig);
  return _charts[canvasId];
}

// ---------------------------------------------------------------------------
// Factory: Line Chart
// ---------------------------------------------------------------------------

export function criarLinha(canvasId, config) {
  destroyIfExists(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;

  const ctx = canvas.getContext('2d');

  const chartConfig = {
    type: 'line',
    data: {
      labels: config.labels || [],
      datasets: config.datasets || [],
    },
    options: baseOptions({
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          max: config.ymax,
          ticks: { callback: (v) => (config.suffix ? v + config.suffix : v) },
        },
      },
      ...config.options,
    }),
    plugins: config.plugins || [],
  };

  _charts[canvasId] = new Chart(ctx, chartConfig);
  return _charts[canvasId];
}

// ---------------------------------------------------------------------------
// Utility: dataset helpers
// ---------------------------------------------------------------------------

export function dataset(label, data, color, extra = {}) {
  return {
    label,
    data,
    backgroundColor: color,
    borderColor: color,
    borderRadius: 4,
    borderSkipped: false,
    ...extra,
  };
}

export { COLORS, STATUS_COLORS };

export function destroyChart(canvasId) {
  destroyIfExists(canvasId);
}
