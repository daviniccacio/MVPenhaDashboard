function mostrarErro(mensagem){
  const modal = document.getElementById("modalErro");
  const texto = document.getElementById("modalErroTexto");
  if(texto) texto.innerText = mensagem;
  modal.classList.remove("hidden");
  modal.classList.add("flex");
}

function fecharModalErro(){
  const modal = document.getElementById("modalErro");
  modal.classList.add("hidden");
  modal.classList.remove("flex");
}

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