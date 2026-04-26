let dados = [];
let chart;

fetch('dados.csv')
  .then(r => r.text())
  .then(t => {
    const linhas = t.split('\n');
    const cabecalho = linhas[0].split(',');

    for (let i = 1; i < linhas.length; i++) {
      if (!linhas[i].trim()) continue;
      const obj = {};
      const valores = linhas[i].split(',');
      cabecalho.forEach((c, j) => obj[c.trim()] = valores[j]?.trim());
      dados.push(obj);
    }

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
  const valores = f.map(d => Number(d.Quantidade) || 0);

  if (chart) chart.destroy();

  const ctx = document.getElementById('grafico');

  chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Quantidade',
        data: valores,
        backgroundColor: '#0078D4'
      }]
    }
  });
}
