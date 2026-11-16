document.addEventListener('DOMContentLoaded', () => {
  const cfg = window.LINTOR_CHART;
  const canvas = document.getElementById('stageChart');
  const selector = document.getElementById('chartType');
  if (!cfg || !canvas || typeof Chart === 'undefined') return;

  const colors = ['#bd2936', '#d65d5f', '#e6a04b', '#5d84ad', '#278766', '#7c6db0'];
  const nf = new Intl.NumberFormat('id-ID');
  Chart.defaults.font.family = "'DM Sans', Arial, sans-serif";
  Chart.defaults.font.size = 11;
  Chart.defaults.color = '#667085';
  let activeChart = null;

  function renderChart(type) {
    if (activeChart) activeChart.destroy();
    const caption = document.getElementById('chartCaption');
    if (type === 'pie') {
      const total = cfg.pie.reduce((sum, value) => sum + (Number(value) || 0), 0);
      activeChart = new Chart(canvas, {
        type: 'pie',
        data: { labels: cfg.stages, datasets: [{ data: cfg.pie, backgroundColor: colors, borderColor: '#fff', borderWidth: 3 }] },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8, padding: 14 } },
            tooltip: { callbacks: { label: c => ` ${c.label}: ${nf.format(c.parsed)} (${total ? (c.parsed / total * 100).toFixed(1).replace('.', ',') : '0,0'}%)` } }
          }
        }
      });
      if (caption) caption.textContent = 'Komposisi akumulasi tiap tahapan pada data yang sedang difilter; bukan persentase terhadap target.';
      return;
    }

    const b = cfg.bar;
    const datasets = Object.keys(b.series).map((name, i) => ({
      label: name, data: b.series[name],
      backgroundColor: name === 'Target' ? 'rgba(29,41,57,.10)' : colors[(i - 1 + colors.length) % colors.length],
      borderColor: name === 'Target' ? '#1d2939' : colors[(i - 1 + colors.length) % colors.length],
      borderWidth: name === 'Target' ? 2 : 0,
      borderDash: name === 'Target' ? [5, 4] : [],
      type: name === 'Target' ? 'line' : 'bar',
      pointRadius: name === 'Target' ? 2 : 0,
      borderRadius: name === 'Target' ? 0 : 3,
      maxBarThickness: 24
    }));
    activeChart = new Chart(canvas, {
      type: 'bar',
      data: { labels: b.labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8, padding: 12 } },
          tooltip: { callbacks: {
            title: items => `${b.labels[items[0].dataIndex]} · Kec. ${b.kecamatan[items[0].dataIndex]}`,
            label: c => ` ${c.dataset.label}: ${c.parsed.y == null ? '—' : nf.format(c.parsed.y)}`
          }}
        },
        scales: {
          x: { grid: { display: false }, ticks: { maxRotation: 55, minRotation: 0, autoSkip: true } },
          y: { beginAtZero: true, grid: { color: '#f0f1f4' }, ticks: { callback: value => nf.format(value) } }
        }
      }
    });
    if (caption) caption.textContent = 'Bandingkan target dan tahapan progres per desa. Maksimal 30 desa ditampilkan agar diagram tetap ringan.';
  }

  renderChart(selector ? selector.value : 'bar');
  if (selector) selector.addEventListener('change', () => renderChart(selector.value));

  // Cascade filters by submitting after parent selection; server options are derived from the same dataset.
  const form = document.getElementById('dashboardFilters');
  const programFilter = document.getElementById('programFilter');
  const kecamatanFilter = document.getElementById('kecamatanFilter');
  const desaFilter = document.getElementById('desaFilter');
  if (form && programFilter) programFilter.addEventListener('change', () => {
    if (kecamatanFilter) kecamatanFilter.value = 'Semua';
    if (desaFilter) desaFilter.value = 'Semua';
    form.requestSubmit();
  });
  if (form && kecamatanFilter) kecamatanFilter.addEventListener('change', () => {
    if (desaFilter) desaFilter.value = 'Semua';
    form.requestSubmit();
  });
});
