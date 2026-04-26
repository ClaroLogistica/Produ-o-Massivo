let dados = [];
let chart = null;
let semanaAtiva = null;
let laboratorioAtivo = null;

/* ===== EXTRAIR DIA ===== */
function extrairDia(data) {
  if (data instanceof Date) return data.getDate();

  if (typeof data === "number") {
    const base = new Date(1899, 11, 30);
    return new Date(base.getTime() + data * 86400000).getDate();
  }

  const d = new Date(data);
  return isNaN(d) ? null : d.getDate();
}

/* ===== CARREGAR EXCEL ===== */
fetch("Dados.xlsx")
  .then(r => r.arrayBuffer())
  .then(b => {
    const wb = XLSX.read(b, { type: "array" });
    const sh = wb.Sheets[wb.SheetNames[0]];
    dados = XLSX.utils.sheet_to_json(sh);

    criarBotoesSemana();
    criarBotoesLaboratorio();

    semanaAtiva = obterSemanas()[0];
    atualizarGrafico();
  });

/* ===== SEMANAS ===== */
function obterSemanas() {
  return [...new Set(dados.map(d => d["semana "]).filter(Boolean))];
}

/* ===== LABORATÓRIOS ===== */
function obterLaboratorios() {
  return [...new Set(dados.map(d => d["Laboratório"]).filter(Boolean))];
}

/* ===== BOTÕES SEMANA ===== */
function criarBotoesSemana() {
  const div = document.getElementById("botoes-semana");
  div.innerHTML = "";

  obterSemanas().forEach((sem, i) => {
    const btn = document.createElement("button");
    btn.textContent = sem;
    btn.className = i === 0 ? "ativo" : "";

    btn.onclick = () => {
      semanaAtiva = sem;
      document.querySelectorAll("#botoes-semana button")
        .forEach(b => b.classList.remove("ativo"));
      btn.classList.add("ativo");
      atualizarGrafico();
    };

    div.appendChild(btn);
  });
}

/* ===== BOTÕES LAB ===== */
function criarBotoesLaboratorio() {
  const div = document.getElementById("botoes-lab");
  div.innerHTML = "";

  const todos = document.createElement("button");
  todos.textContent = "Todos";
  todos.className = "ativo";
  todos.onclick = () => {
    laboratorioAtivo = null;
    document.querySelectorAll("#botoes-lab button")
      .forEach(b => b.classList.remove("ativo"));
    todos.classList.add("ativo");
    atualizarGrafico();
  };
  div.appendChild(todos);

  obterLaboratorios().forEach(lab => {
    const btn = document.createElement("button");
    btn.textContent = lab;

    btn.onclick = () => {
      laboratorioAtivo = lab;
      document.querySelectorAll("#botoes-lab button")
        .forEach(b => b.classList.remove("ativo"));
      btn.classList.add("ativo");
      atualizarGrafico();
    };

    div.appendChild(btn);
  });
}

/* ===== GRÁFICO ===== */
function atualizarGrafico() {
  const labels = Array.from({ length: 31 }, (_, i) => i + 1);
  const valores = Array(31).fill(0);

  dados
    .filter(d => d["semana "] === semanaAtiva)
    .filter(d => !laboratorioAtivo || d["Laboratório"] === laboratorioAtivo)
    .forEach(d => {
      const dia = extrairDia(d.Data);
      if (dia) valores[dia - 1] += Number(d.Quantidade || 0);
    });

  if (chart) chart.destroy();

  chart = new Chart(document.getElementById("graficoDiario"), {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "Produção por Dia",
        data: valores,
        backgroundColor: "#38bdf8",
        barThickness: 12
      }]
    },
    options: {
      animation: false,
      scales: {
        x: { ticks: { color: "#e5e7eb" } },
        y: { beginAtZero: true, ticks: { color: "#e5e7eb" } }
      },
      plugins: {
        legend: { labels: { color: "#e5e7eb" } }
      }
    },
    plugins: [{
      id: "valoresTopo",
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        ctx.save();
        ctx.fillStyle = "#e5e7eb";
        ctx.font = "11px Arial";
        ctx.textAlign = "center";

        chart.getDatasetMeta(0).data.forEach((bar, i) => {
          const valor = chart.data.datasets[0].data[i];
          if (valor > 0) {
            ctx.fillText(valor.toLocaleString("pt-BR"), bar.x, bar.y - 5);
          }
        });

        ctx.restore();
      }
    }]
  });
}
