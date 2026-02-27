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
                const emojis = {
                    "GASOLINA": "⛽",
                    "DIESEL": "🚛",
                    "DIESEL S10": "🚚",
                    "OUTROS": "🚗"
                }
                const emoji = emojis[ctx.label?.toUpperCase()] || "🔋";
                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                const pct   = ((ctx.parsed / total) * 100).toFixed(1);
                return `${emoji}    ${ctx.label}: ${ctx.parsed.toLocaleString("pt-BR")} L (${pct}%)`;
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
