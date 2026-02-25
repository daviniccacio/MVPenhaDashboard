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

function processarDados(dadosRaw) {
  let dados = dadosRaw[0] && normalizarItem(dadosRaw[0])["Placa"]?.toString().trim() === "Placa" ? dadosRaw.slice(1) : dadosRaw;

  if(dados[0] && normalizarItem(dados[0])["Placa"]?.toString().trim() === "Placa"){
    dados = dados.slice(1);
  };

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

    const placaRaw = item["Placa"]?.toString().trim().split(/\s+/) || [];
    const placa = placaRaw[0] || "SEM PLACA";
    const comb = item["Combustível"]?.toString().trim() || placaRaw.slice(1).join(" ") || "OUTRO";

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

  document.getElementById("totalRegistros").innerText = dados.filter(item => {
    const i = normalizarItem(item);
    return i["Placa"] && !i["Placa"].toString().toLowerCase().includes("total") && i["Placa"].toString().trim().toLowerCase() !== "placa";
  }).length;

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

  const btnExportar = document.getElementById("btnExportar");
  btnExportar.disabled = false;
  btnExportar.classList.add("bg-blue-700", "hover:bg-blue-600", "cursor-pointer");
  btnExportar.classList.remove("bg-slate-700");
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
    btnGraf.classList.add("ring-2", "ring-white");
    btnTab.classList.remove("ring-2", "ring-white");
  } else {
    tabela.classList.remove("hidden");
    graficos.classList.add("hidden");
    btnTab.classList.add("bg-blue-600", "text-white");
    btnTab.classList.remove("bg-slate-700", "text-slate-300");
    btnGraf.classList.add("bg-slate-700", "text-slate-300");
    btnGraf.classList.remove("bg-blue-600", "text-white");
    btnTab.classList.add("ring-2", "ring-white");
    btnGraf.classList.remove("ring-2", "ring-white");
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
    const placaRaw = (item["Placa"]?.toString().trim() || "").split(/\s+/);
    const placa = placaRaw[0] || "SEM PLACA";
    const comb = item["Combustível"]?.toString().trim() || placaRaw.slice(1).join(" ") || "OUTRO";
    const litrosAutorizados = limparNum(item["Autorizado"]);
    const litrosEfetivado = limparNum(item["Efetivado"]);
    const vlrTotalDireto = limparNum(item["Vlr Total"]);
    const vlrUnitario    = limparNum(item["Vlr Unitário"]);
    const valor  = vlrTotalDireto || (litrosEfetivado * vlrUnitario);
    const data   = excelParaData(item["Data/Hora"]);
    const motoristaRaw = item["Motorista"]?.toString().trim();
    const motorista = motoristaRaw || item["Vlr Total"]?.toString().split(/\s{2,}/)[1]?.trim() || "-";
    

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

async function exportarPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

  const fileName    = document.getElementById("fileName").innerText;
  const totalGeral  = document.getElementById("totalGeral").innerText;
  const totalLitros = document.getElementById("totalLitros").innerText;
  const totalReg    = document.getElementById("totalRegistros").innerText;
  const media       = document.getElementById("mediaUnit").innerText;

  // ---- PÁGINA 1 — CARDS + GRÁFICOS ----
  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, 297, 210, "F");

  doc.setFontSize(16);
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.text("MV Penha Dashboard", 14, 14);

  doc.setFontSize(8);
  doc.setTextColor(148, 163, 184);
  doc.setFont("helvetica", "normal");
  doc.text(`Planilha: ${fileName}  |  Gerado em: ${new Date().toLocaleDateString("pt-BR")}`, 14, 20);

  // Cards de Resumo
  const cards = [
    { label: "Valor Total",    value: totalGeral,  color: [16, 185, 129] },
    { label: "Volume Total",   value: totalLitros, color: [59, 130, 246] },
    { label: "Registros",      value: totalReg,    color: [168, 85, 247] },
    { label: "Média Unitária", value: media,       color: [251, 146, 60] },
  ];

  cards.forEach((card, i) => {
    const x = 14 + i * 70;
    doc.setFillColor(30, 41, 59);
    doc.roundedRect(x, 24, 64, 18, 2, 2, "F");
    doc.setTextColor(148, 163, 184);
    doc.setFontSize(7);
    doc.setFont("helvetica", "normal");
    doc.text(card.label, x + 4, 30);
    doc.setTextColor(...card.color);
    doc.setFontSize(10);
    doc.setFont("helvetica", "bold");
    doc.text(card.value, x + 4, 37);
  });

  // Gráficos — 2x2
  const graficos = [
    { id: "chartCombustivel",     x: 14,  y: 46 },
    { id: "chartEstabelecimento", x: 154, y: 46 },
    { id: "chartVeiculos",        x: 14,  y: 122 },
    { id: "chartComparativo",     x: 154, y: 122 },
  ];

  for (const g of graficos) {
    const canvas = document.getElementById(g.id);
    if (canvas) {
      doc.setFillColor(30, 41, 59);
      doc.roundedRect(g.x, g.y, 130, 72, 2, 2, "F");

      const img = canvas.toDataURL("image/png", 1.0);
      
      // Cálculo de proporção para caber no box
      const imgProps = doc.getImageProperties(img);
      const margin = 5;
      const maxW = 120;
      const maxH = 62;
      const ratio = imgProps.width / imgProps.height;
      
      let w = maxW;
      let h = w / ratio;
      if (h > maxH) {
        h = maxH;
        w = h * ratio;
      }

      const centerOffX = (130 - w) / 2;
      const centerOffY = (72 - h) / 2;

      // CORREÇÃO: g.y em vez de g.x para o Offset vertical
      doc.addImage(img, "PNG", g.x + centerOffX, g.y + centerOffY, w, h);
    }
  }

  // ---- PÁGINA 2 — TABELA ----
  doc.addPage();
  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, 297, 210, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(13);
  doc.setFont("helvetica", "bold");
  doc.text("Detalhamento de Abastecimentos", 14, 13);

  doc.setFontSize(8);
  doc.setTextColor(148, 163, 184);
  doc.setFont("helvetica", "normal");
  doc.text(`Planilha: ${fileName}`, 14, 19);

  // Coleta linhas da tabela HTML
  const linhas = [];
  document.querySelectorAll("#tabelaCorpo tr").forEach((tr) => {
    const cols = tr.querySelectorAll("td");
    if (cols.length > 0) {
      linhas.push([
        cols[0].innerText,
        cols[1].innerText,
        cols[2].innerText,
        cols[3].innerText,
        cols[4].innerText,
        cols[5].innerText,
        cols[6].innerText,
        cols[7].innerText,
        cols[8].innerText,
      ]);
    }
  });

  doc.autoTable({
    startY: 23,
    head: [["Data", "Placa", "Estabelecimento", "Combustível", "Litros Autorizados", "Litros Efetivados", "Valor Unitário", "Valor Total", "Motorista"]],
    body: linhas,
    theme: "grid",
    styles: {
      fillColor: [30, 41, 59],
      textColor: [226, 232, 240],
      fontSize: 8,
      cellPadding: 3,
    },
    headStyles: {
      fillColor: [51, 65, 85],
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: {
      fillColor: [15, 23, 42],
    },
    columnStyles: {
      0: { cellWidth: 22 }, // Data
      1: { cellWidth: 20 }, // Placa
      2: { cellWidth: 55 }, // Estabelecimento
      3: { cellWidth: 28 }, // Combustível
      4: { cellWidth: 25 }, // Litros Autorizados
      5: { cellWidth: 25 }, // Litros Efetivados
      6: { cellWidth: 22 }, // Valor Unitário
      7: { cellWidth: 22 }, // Valor Total
      8: { cellWidth: 48 }, // Motorista
    },
  });

  const pdfBlob = doc.output("blob");
  const pdfUrl = URL.createObjectURL(pdfBlob);

  window._pdfUrl = pdfUrl;
  window._pdfNome = `RELATORIO ${fileName.replace(".xlsx", "")}.pdf`;

  mostrarModalPDF();
}

function mostrarModalPDF(){
  const modal = document.getElementById("modalPDF");
  modal.classList.remove("hidden");
  modal.classList.add("flex");
}

function fecharModalPDF(){
  const modal = document.getElementById("modalPDF");
  modal.classList.add("hidden");
  modal.classList.remove("flex");
}

function baixarPDF(){
  const a = document.createElement("a");
  a.href = window._pdfUrl;
  a.download = window._pdfNome;
  a.click();
  fecharModalPDF();
}

function abrirPDF(){
  window.open(window._pdfUrl, "_blank");
  fecharModalPDF();
}