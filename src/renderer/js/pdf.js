async function exportarPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

  const fileName = document.getElementById("fileName").innerText;
  const totalGeral = document.getElementById("totalGeral").innerText;
  const totalLitros = document.getElementById("totalLitros").innerText;
  const totalReg = document.getElementById("totalRegistros").innerText;
  const media = document.getElementById("mediaUnit").innerText;

  // Função para desenhar o fundo e o cabeçalho (usada na pag 1 e nas novas)
  const desenharLayoutBase = (docInstance) => {
    docInstance.setFillColor(15, 23, 42);
    docInstance.rect(0, 0, 297, 210, "F");
  };

  // ---- PÁGINA 1 — CARDS + GRÁFICOS ----
  desenharLayoutBase(doc);

  doc.setFontSize(16);
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.text("MV Penha Dashboard", 14, 14);

  doc.setFontSize(8);
  doc.setTextColor(148, 163, 184);
  doc.text(`Planilha: ${fileName}  |  Gerado em: ${new Date().toLocaleDateString("pt-BR")}`, 14, 20);

  // Cards de Resumo
  const cards = [
    { label: "Valor Total", value: totalGeral, color: [16, 185, 129] },
    { label: "Volume Total", value: totalLitros, color: [59, 130, 246] },
    { label: "Registros", value: totalReg, color: [168, 85, 247] },
    { label: "Média Unitária", value: media, color: [251, 146, 60] },
  ];

  cards.forEach((card, i) => {
    const x = 14 + i * 70;
    doc.setFillColor(30, 41, 59);
    doc.roundedRect(x, 24, 66, 18, 2, 2, "F");
    doc.setTextColor(148, 163, 184);
    doc.setFontSize(7);
    doc.text(card.label, x + 4, 30);
    doc.setTextColor(...card.color);
    doc.setFontSize(10);
    doc.setFont("helvetica", "bold");
    doc.text(card.value, x + 4, 37);
  });

  // Gráficos
  const graficos = [
    { id: "chartCombustivel", x: 14, y: 46 },
    { id: "chartEstabelecimento", x: 154, y: 46 },
    { id: "chartVeiculos", x: 14, y: 122 },
    { id: "chartComparativo", x: 154, y: 122 },
  ];

  for (const g of graficos) {
    const canvas = document.getElementById(g.id);
    if (canvas) {
      const tempCanvas = document.createElement('canvas');
      tempCanvas.width = canvas.width;
      tempCanvas.height = canvas.height;
      const ctx = tempCanvas.getContext('2d');

      ctx.fillStyle = '#1e293b';
      ctx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);

      ctx.drawImage(canvas, 0, 0);

      const img = tempCanvas.toDataURL("image/png", 1.0);

      doc.setFillColor(30, 41, 59);
      doc.roundedRect(g.x, g.y, 130, 72, 2, 2, "F");

      const maxW = 124;
      const maxH = 66;
      doc.addImage(img, "PNG", g.x + 3, g.y + 3, maxW, maxH);
    }
  }

  // ---- PÁGINA 2 — TABELA ----
  doc.addPage();
  // O layout da página 2 e seguintes será controlado pelo autoTable

  const linhas = [];
  document.querySelectorAll("#tabelaCorpo tr").forEach((tr) => {
    const cols = tr.querySelectorAll("td");
    if (cols.length > 0) {
      linhas.push(Array.from(cols).map(td => td.innerText));
    }
  });

  doc.autoTable({
    startY: 25,
    head: [["Data", "Placa", "Estabelecimento", "Combustível", "L. Aut.", "L. Efet.", "Vlr. Unit.", "Vlr. Total", "Motorista"]],
    body: linhas,
    theme: "grid",
    // 1. WILL DRAW PAGE: Pinta o fundo ANTES da tabela aparecer
    willDrawPage: function(data) {
      desenharLayoutBase(doc);
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(13);
      doc.setFont("helvetica", "bold");
      doc.text("Detalhamento de Abastecimentos", 14, 15);
    },
    styles: {
      fillColor: [30, 41, 59],
      textColor: [226, 232, 240],
      fontSize: 8,
      cellPadding: 2,
    },
    headStyles: {
      fillColor: [51, 65, 85],
      textColor: [255, 255, 255],
    },
    alternateRowStyles: {
      fillColor: [15, 23, 42],
    },
    // Ajuste de larguras para não estourar a página (Total < 265mm)
    columnStyles: {
      0: { cellWidth: 20 }, // Data
      1: { cellWidth: 18 }, // Placa
      2: { cellWidth: 50 }, // Estabelecimento
      3: { cellWidth: 25 }, // Combustível
      4: { cellWidth: 22 }, // L. Aut
      5: { cellWidth: 22 }, // L. Efet
      6: { cellWidth: 22 }, // Vlr Unit
      7: { cellWidth: 22 }, // Vlr Total
      8: { cellWidth: 'auto' }, // Motorista ocupa o que sobrar
    },
    margin: { left: 14, right: 14, top: 25 }
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