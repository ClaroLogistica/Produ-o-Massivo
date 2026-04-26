let dados = [];
let chart;

fetch('Dados.xlsx')
  .then(r => r.arrayBuffer())
  .then(b => {
    const wb = XLSX.read(b, { type: 'array' });
    const sh = wb.Sheets[wb.SheetNames[0]];
    dados = XLSX.utils.sheet_to_json(sh);

    criarFiltros();
    atualizar();
  });

function criarFiltros() {
  document.querySelectorAll('select').forEach(sel => {
    const col = sel.dataset.col;
    const valores = [...new Set(dados.map(d => d[col]).filter(v => v))];
    sel.innerHTML = `<option value="">Todos</option>`;
    valores.forEach(v => sel.innerHTML += `<option>${v}</option>`);
    sel.onchange = atualizar;
  });
}

function atualizar() {
  let f = [...dados];

  document.querySelectorAll('select').forEach(sel => {
    if (sel.value) {
      f = f.filter(d => String(d[sel.dataset.col]) === sel.value);
    }
  });

  // ✅ AGRUPAR PRODUÇÃO (somar Quantidade)
  const agrupado = {};
  f.forEach(d => {
    const chave = d['Tipo'] || 'Sem Tipo';
    agrupado[chave] = (agrupado[chave] || 0) + Number(d['Quantidade'] || 0);
  });

  const labels = Object.keys(agrupado);
  const valores = Object.values(agrupado);

  if (chart) chart.destroy();

  chart = new Chart(document.getElementById('grafico'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Quantidade Produzida',
        data: valores,
        backgroundColor: '#0078D4'
      }]
    },
    options: {
      responsive: true,
      scales: {
        y: { beginAtZero: true }
      }
    }
  });
}
