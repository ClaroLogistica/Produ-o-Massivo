let dados = [];
let chart = null;
let localAtivo = null;
let terminalAtivo = null;

function extrairDia(data) {
  if (data instanceof Date) return data.getDate();
  if (typeof data === "number") {
    const base = new Date(1899, 11, 30);
    return new Date(base.getTime() + data * 86400000).getDate();
  }
  const d = new Date(data);
  return isNaN(d) ? null : d.getDate();
}

fetch("Dados.xlsx")
  .then(r => r.arrayBuffer())
  .then(b => {
    const wb = XLSX.read(b, { type: "array" });
    const sh = wb.Sheets[wb.SheetNames[0]];
    dados = XLSX.utils.sheet_to_json(sh);

    criarBotoesLocal();
    criarBotoesTerminais();
    atualizarTudo();
  });

function obterLocais() {
  return [...new Set(dados.map(d => d["Local"]).filter(Boolean))];
}

function obterTerminais() {
  return [...new Set(dados.map(d => d["Terminais"]).filter(Boolean))];
}

function criarBotoesLocal() {
  const div = document.getElementById("botoes-lab");
  div.innerHTML = "";

  const todos = document.createElement("button");
  todos.textContent = "Todos";
  todos.classList.add("ativo");
  todos.onclick = () => { localAtivo = null; atualizarTudo(); };
  div.appendChild(todos);

  obterLocais().forEach(l => {
    const b = document.createElement("button");
    b.textContent = l;
    b.onclick = () => { localAtivo = l; atualizarTudo(); };
    div.appendChild(b);
  });
}

function criarBotoesTerminais() {
  const div = document.getElementById("botoes-terminais");
  div.innerHTML = "";

  const todos = document.createElement("button");
  todos.textContent = "Todos";
  todos.classList.add("ativo");
  todos.onclick = () => { terminalAtivo = null; atualizarTudo(); };
  div.appendChild(todos);

  obterTerminais().forEach(t => {
    const b = document.createElement("button");
    b.textContent = t;
    b.onclick = () => { terminalAtivo = t; atualizarTudo(); };
    div.appendChild(b);
  });
}

function atualizarTudo() {
  atualizarKPIs();
  atualizarGrafico();
  atualizarResumoSemanal();
}

function atualizarKPIs() {
  const base = dados
    .filter(d => !localAtivo || d["Local"] === localAtivo)
    .filter(d => !terminalAtivo || d["Terminais"] === terminalAtivo);

  const totalSelecionado = base.reduce((s, d) => s + Number(d.Quantidade || 0), 0);
  const totalMes = dados.reduce((s, d) => s + Number(d.Quantidade || 0), 0);

  document.getElementById("kpi-selecionado").textContent = totalSelecionado.toLocaleString("pt-BR");
  document.getElementById("kpi-mes").textContent = totalMes.toLocaleString("pt-BR");
}

function atualizarGrafico() {
  const labels = Array.from({ length: 31 }, (_, i) => i + 1);
  const valores = Array(31).fill(0);

  dados
    .filter(d => !localAtivo || d["Local"] === localAtivo)
    .filter(d => !terminalAtivo || d["Terminais"] === terminalAtivo)
    .forEach(d => {
      const dia = extrairDia(d.Data);
      if (dia) valores[dia - 1] += Number(d.Quantidade || 0);
    });

  if (chart) chart.destroy();
  chart = new Chart(document.getElementById("graficoDiario"), {
    type: "bar",
    data: { labels, datasets: [{ data: valores, backgroundColor: "#38bdf8" }] }
  });
}

function atualizarResumoSemanal() {
  const c = document.getElementById("resumo-semanal");
  c.innerHTML = "";

  const base = dados
    .filter(d => !localAtivo || d["Local"] === localAtivo)
    .filter(d => !terminalAtivo || d["Terminais"] === terminalAtivo);

  const totalMes = base.reduce((s, d) => s + Number(d.Quantidade || 0), 0);

  const porSemana = {};
  base.forEach(d => {
    const s = d["Semana"];
    if (!s) return;
    porSemana[s] = (porSemana[s] || 0) + Number(d.Quantidade || 0);
  });

  Object.keys(porSemana).sort().forEach(s => {
    const total = porSemana[s];
    const p = totalMes ? Math.round((total / totalMes) * 100) : 0;
    const div = document.createElement("div");
    div.className = "sem-box";
    div.innerHTML = `<span>${s}</span><span>${total.toLocaleString("pt-BR")}</span><span class="percentual">${p}%</span>`;
    c.appendChild(div);
  });
}
