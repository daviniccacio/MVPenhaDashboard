let charts = {};

document.addEventListener("DOMContentLoaded", () => {
  const btnLoad = document.getElementById("btnLoad");

  btnLoad.addEventListener("click", async () => {
    const response = await window.electronAPI.openExcel();
    if (!response) return;

    const { data, fileName } = response;
    document.getElementById("fileName").innerText = fileName;
    processarDados(data);
  });
});

function normalizarItem(itemRaw) {
  const mapa = {
    "combustivel":          "Combustível",
    "combustível":          "Combustível",
    "vir total motorista":  "Vlr Total",
    "vlr total motorista":  "Vlr Total",
    "vlr total":            "Vlr Total",
    "vir combust.":         "Vlr Unitário",
    "vlr unitário":         "Vlr Unitário",
    "vlr unitario":         "Vlr Unitário",
    "hora":                 "Data/Hora",
  };

  return Object.fromEntries(
    Object.entries(itemRaw).map(([k, v]) => {
      const chaveNorm = mapa[k.trim().toLowerCase()] || k.trim();
      return [chaveNorm, v];
    })
  );
}

function processarDados(dados) {
  let totalGeral       = 0;
  let totalLitros      = 0;
  let combustiveis     = {};
  let estabelecimentos = {};
  let veiculos         = {};
  let autorizado       = 0;
  let efetivado        = 0;

  const limparNum = (val) => {
    if (!val) return 0;
    // Se já é número (xlsx converte automaticamente), retorna direto
    if (typeof val === "number") return val;
    // Remove motorista que vem junto ex: "1.158,00  MARCOS"
    const str = val.toString().split(/\s{2,}/)[0].trim();
    // Remove pontos de milhar, troca vírgula por ponto decimal
    return parseFloat(str.replace(/\./g, "").replace(",", ".")) || 0;
  };

  dados.forEach((itemRaw) => {
    const item = normalizarItem(itemRaw);

    // Pula linhas sem placa ou linha de totais
    if (!item["Placa"]) return;
    if (item["Placa"].toString().trim().toLowerCase() === "placa") return;
    if (item["Placa"].toString().toLowerCase().includes("total")) return;

    // Limpa nome do estabelecimento
    const estBruto = item["Estabelecimento"]?.toString() || "NÃO INFORMADO";
    const est = estBruto.replace(/Razão Social:/gi, "").trim().toUpperCase();

    const comb  = item["Combustível"]?.toString().trim() || "OUTRO";
    const placa = item["Placa"]?.toString().trim()       || "SEM PLACA";

    const litros         = limparNum(item["Efetivado"]);
    const vlrUnitario    = limparNum(item["Vlr Unitário"]);
    const vlrAutorizado  = limparNum(item["Autorizado"]);
    const vlrTotalDireto = limparNum(item["Vlr Total"]);

    // Usa Vlr Total direto, senão calcula litros × preço unitário
    const vlrFinal = vlrTotalDireto || (litros * vlrUnitario);

    totalGeral  += vlrFinal;
    totalLitros += litros;
    autorizado  += vlrAutorizado;
    efetivado   += litros;

    if (litros > 0) {
      combustiveis[comb] = (combustiveis[comb] || 0) + litros;
    }
    if (vlrFinal > 0) {
      estabelecimentos[est] = (estabelecimentos[est] || 0) + vlrFinal;
      veiculos[placa]       = (veiculos[placa]       || 0) + vlrFinal;
    }
  });

  console.log("Total de linhas:", dados.length);
  console.log("Primeira linha:", JSON.stringify(normalizarItem(dados[0])));
  console.log("Segunda linha:", JSON.stringify(normalizarItem(dados[1])));

  // Atualiza os cards
  document.getElementById("totalGeral").innerText =
    totalGeral.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

  document.getElementById("totalLitros").innerText =
    `${totalLitros.toLocaleString("pt-BR", { minimumFractionDigits: 0 })} L`;

  document.getElementById("totalRegistros").innerText = dados.length;

  const media = totalLitros > 0 ? totalGeral / totalLitros : 0;
  document.getElementById("mediaUnit").innerText =
    media.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

  Chart.defaults.color       = "#e2e8f0";
  Chart.defaults.borderColor = "rgba(255, 255, 255, 0.1)";
  Chart.defaults.font.family = "'Inter'";

  renderPizza("chartCombustivel", combustiveis);
  renderBarras("chartEstabelecimento", estabelecimentos, "Gasto por Estabelecimento (R$)");
  renderBarras("chartVeiculos", veiculos, "Gasto por Veículo (R$)", true, 10);
  renderComparativo("chartComparativo", autorizado, efetivado);

  renderTabela(dados);
}

// Controla qual aba está visível
function trocarAba(aba) {
  const graficos = document.getElementById("secaoGraficos");
  const tabela   = document.getElementById("secaoTabela");
  const btnGraf  = document.getElementById("tabGraficos");
  const btnTab   = document.getElementById("tabTabela");

  if (aba === "graficos") {
    graficos.classList.remove("hidden");
    tabela.classList.add("hidden");
    btnGraf.classList.add("bg-blue-600", "text-white");
    btnGraf.classList.remove("bg-slate-700", "text-slate-300");
    btnTab.classList.add("bg-slate-700", "text-slate-300");
    btnTab.classList.remove("bg-blue-600", "text-white");
  } else {
    tabela.classList.remove("hidden");
    graficos.classList.add("hidden");
    btnTab.classList.add("bg-blue-600", "text-white");
    btnTab.classList.remove("bg-slate-700", "text-slate-300");
    btnGraf.classList.add("bg-slate-700", "text-slate-300");
    btnGraf.classList.remove("bg-blue-600", "text-white");
  }
}

// Preenche a tabela com os dados
function renderTabela(dados) {
  const corpo = document.getElementById("tabelaCorpo");
  corpo.innerHTML = "";

  const limparNum = (val) => {
    if (!val) return 0;
    if (typeof val === "number") return val;
    const str = val.toString().split(/\s{2,}/)[0].trim();
    return parseFloat(str.replace(/\./g, "").replace(",", ".")) || 0;
  };

  // Converte número serial do Excel para data
  const excelParaData = (val) => {
    if (!val) return "-";
    if (typeof val === "number") {
      // Excel conta dias desde 01/01/1900
      const data = new Date((val - 25569) * 86400 * 1000);
      return data.toLocaleDateString("pt-BR");
    }
    if (val instanceof Date) return val.toLocaleDateString("pt-BR");
    return val.toString().slice(0, 10);
  };

  dados.forEach((itemRaw) => {
    const item = normalizarItem(itemRaw);
    if (!item["Placa"]) return;
    if (item["Placa"].toString().toLowerCase().includes("total")) return;
    if (item["Placa"].toString().trim().toLowerCase() === "placa") return;

    const est    = item["Estabelecimento"]?.toString().replace(/Razão Social:/gi, "").trim() || "-";
    const placa  = item["Placa"]?.toString().trim() || "-";
    const comb   = item["Combustível"]?.toString().trim() || "-";
    const litrosAutorizados = limparNum(item["Autorizado"]);
    const litrosEfetivado = limparNum(item["Efetivado"]);
    const vlrTotalDireto = limparNum(item["Vlr Total"]);
    const vlrUnitario    = limparNum(item["Vlr Unitário"]);
    const valor  = vlrTotalDireto || (litros * vlrUnitario);
    const data   = excelParaData(item["Data/Hora"]);
    const motorista = item["Motorista"].toString().trim() || "-";

    corpo.innerHTML += `
      <tr class="border-b border-slate-700/50 hover:bg-slate-700/30 transition-colors">
        <td class="px-4 py-3 text-slate-300">${data}</td>
        <td class="px-4 py-3 font-mono font-bold text-blue-400">${placa}</td>
        <td class="px-4 py-3 text-slate-300">${est}</td>
        <td class="px-4 py-3 text-slate-300">${comb}</td>
        <td class="px-4 py-3 text-slate-300">${litrosAutorizados.toLocaleString("pt-BR")} L</td>
        <td class="px-4 py-3 text-slate-300">${litrosEfetivado.toLocaleString("pt-BR")} L</td>
        <td class="px-4 py-3 text-emerald-400 font-semibold">${vlrUnitario.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</td>
        <td class="px-4 py-3 text-emerald-400 font-semibold">${valor.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</td>
        <td class="px-4 py-3 text-slate-300">${motorista}</td>
      </tr>
    `;
  });
}

function renderPizza(canvasId, dados) {
  if (charts[canvasId]) charts[canvasId].destroy();
  const ctx = document.getElementById(canvasId)?.getContext("2d");
  if (!ctx) return;

  charts[canvasId] = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: Object.keys(dados),
      datasets: [{
        borderWidth: 0,
        data: Object.values(dados),
        backgroundColor: ["#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6"],
      }],
    },
    options: {
      cutout: "50%",
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 1200, easing: "easeOutQuart", animateRotate: true, animateScale: true },
      plugins: {
        legend: { position: "bottom", labels: { usePointStyle: true, padding: 20 } },
        tooltip: {
          backgroundColor: "#1e293b",
          titleColor: "#fff",
          titleAlign: "center",
          bodyColor: "#cbd5e1",
          padding: 12,
          boxPadding: 8,
          borderColor: "#334155",
          borderWidth: 1,
          callbacks: {
            label: (ctx) => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct   = ((ctx.parsed / total) * 100).toFixed(1);
              return `${ctx.label}: ${ctx.parsed.toLocaleString("pt-BR")} L (${pct}%)`;
            },
          },
        },
      },
    },
  });
}

function renderBarras(canvasId, dados, label, horizontal = false, limite = 8) {
  if (charts[canvasId]) charts[canvasId].destroy();
  const ctx = document.getElementById(canvasId)?.getContext("2d");
  if (!ctx) return;

  const quebrarTexto = (text, maxChars) => {
    const words = text.split(" ");
    let lines = [];
    let currentLine = "";
    words.forEach((word) => {
      if ((currentLine + word).length > maxChars) {
        lines.push(currentLine.trim());
        currentLine = word + " ";
      } else {
        currentLine += word + " ";
      }
    });
    lines.push(currentLine.trim());
    return lines;
  };

  const sorted = Object.entries(dados)
    .sort(([, a], [, b]) => b - a)
    .slice(0, limite);

  charts[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: sorted.map(([k]) => quebrarTexto(k, 12)),
      datasets: [{
        label,
        data: sorted.map(([, v]) => v),
        backgroundColor: "#6D73C9",
        borderRadius: 5,
        barThickness: 20,
      }],
    },
    options: {
      indexAxis: horizontal ? "y" : "x",
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 1000, easing: "easeOutQuart" },
      scales: {
        y: { beginAtZero: true, grid: { color: "rgba(255, 255, 255, 0.05)" } },
        x: {
          grid: { display: false },
          ticks: { autoSkip: false, maxRotation: 0, minRotation: 0, font: { size: 10 }, padding: 5 },
        },
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "#1e293b",
          titleColor: "#fff",
          titleAlign: "center",
          bodyColor: "#cbd5e1",
          padding: 12,
          boxPadding: 8,
          borderColor: "#334155",
          borderWidth: 1,
          callbacks: {
            label: (ctx) => {
              const val = horizontal ? ctx.parsed.x : ctx.parsed.y;
              return val.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
            },
          },
        },
      },
    },
  });
}

function renderComparativo(canvasId, autorizado, efetivado) {
  if (charts[canvasId]) charts[canvasId].destroy();
  const ctx = document.getElementById(canvasId)?.getContext("2d");
  if (!ctx) return;

  charts[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ["Comparativo"],
      datasets: [
        { label: "Autorizado", data: [autorizado], backgroundColor: "#f59e0b", borderRadius: 6 },
        { label: "Efetivado",  data: [efetivado],  backgroundColor: "#10b981", borderRadius: 6 },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 1000, easing: "easeOutQuart" },
      plugins: {
        legend: { labels: { color: "#f1f5f9" } },
        tooltip: {
          backgroundColor: "#1e293b",
          titleColor: "#fff",
          titleAlign: "center",
          bodyColor: "#cbd5e1",
          padding: 12,
          boxPadding: 8,
          borderColor: "#334155",
          borderWidth: 1,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString("pt-BR")} L`,
          },
        },
      },
      scales: {
        x: { ticks: { color: "#94a3b8" }, grid: { display: false } },
        y: { beginAtZero: true, ticks: { padding: 10 }, grid: { color: "rgba(255,255,255,0.05)" } },
      },
    },
  });
}