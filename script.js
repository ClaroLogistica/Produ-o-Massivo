let dados = [];
let chart;

fetch('dados.xlsx')
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
    if (sel.value) f = f.filter(d => d[sel.dataset.col] == sel.value);
  });

  const labels = f.map(d => d.Descrição || d.Material);
  const valores = f.map(d => d.Quantidade);

  if (chart) chart.destroy();

  chart = new Chart(grafico, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: 'Quantidade', data: valores }]
    }
  });
}
