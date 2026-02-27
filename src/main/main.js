const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const os = require('os');
const { Menu } = require('electron');
const path = require("path");
const xlsx = require("xlsx");
const fs = require('fs');

// Menu.setApplicationMenu(null);

const historicoPath = path.join(os.homedir(), ".mvpenha_historico.json");

function lerHistorico(){
  try{
    if(fs.existsSync(historicoPath)){
      return JSON.parse(fs.readFileSync(historicoPath, "utf-8"));
    }
  } catch {}
  return [];
}

function salvarHistorico(fileName, filePath){
  let historico = lerHistorico();
  historico = historico.filter(r => r.path !== filePath);
  historico.unshift({ name: fileName, path: filePath, data: new Date().toLocaleDateString("pt-BR") });
  historico = historico.slice(0, 3);
  fs.writeFileSync(historicoPath, JSON.stringify(historico, null, 2));
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    icon: path.join(__dirname, "../../assets/icon/mvpenhalogo2.ico"),
    webPreferences: {
      preload: path.join(__dirname, "../preload/preload.js"),
      nodeIntegration: false,
      contextIsolation: true,
    },
  });
  win.loadFile("src/renderer/index.html");
}

ipcMain.handle("open-excel", async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Planilhas", extensions: ["xlsx", "xls"] }],
  });

  if (canceled) return null;

  try {
    const workbook = xlsx.readFile(filePaths[0]);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Lê todas as linhas como array puro para detectar o cabeçalho
    const jsonRaw = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Detecta a linha que tem "Placa" E "Estabelecimento" juntos
    let headerRow = 0;
    for (let i = 0; i < jsonRaw.length; i++) {
      const row = jsonRaw[i];
      const temPlaca = row.some((c) => c?.toString().trim() === "Placa");
      const temEst   = row.some((c) => c?.toString().trim() === "Estabelecimento");
      if (temPlaca && temEst) {
        headerRow = i;
        break;
      }
    }

    console.log("Header detectado na linha:", headerRow);
    console.log("Conteúdo:", jsonRaw[headerRow]);

    // Pega os nomes reais das colunas da linha de cabeçalho (ignora células vazias)
    const headers = jsonRaw[headerRow];

    // Lê os dados a partir da linha de cabeçalho
    const rawData = xlsx.utils.sheet_to_json(sheet, {
      range: headerRow,
      defval: "",
      blankrows: false,
    });

    // Corrige colunas que vieram como __EMPTY usando os nomes reais do cabeçalho
    const cleanData = rawData.map((row) => {
      const entries = Object.entries(row);
      const result  = {};

      entries.forEach(([k, v], i) => {
        if (k.startsWith("__EMPTY")) {
          // Usa o nome real da coluna pelo índice
          const nomeReal = headers[i]?.toString().trim() || k;
          result[nomeReal] = v;
        } else {
          result[k.trim()] = v;
        }
      });

      return result;
    });

    salvarHistorico(path.basename(filePaths[0]), filePaths[0]);

    return {
      data: cleanData,
      fileName: path.basename(filePaths[0]),
    };
  } catch (error) {
    console.error("Erro ao ler Excel:", error);
    return null;
  }
});

ipcMain.handle("get-historico", () => lerHistorico());

ipcMain.handle("open-excel-path", async (event, filePath) => {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const jsonRaw = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    let headerRow = 0;
    for(let i = 0;i < jsonRaw.length;i++){
      const row = jsonRaw[i];
      const temPlaca = row.some(c => c?.toString().trim() === "Placa");
      const temEst = row.some(c => c?.toString().trim() === "Estabelecimento");
      if(temPlaca && temEst) { headerRow = i; break; }
    }

    const headers = jsonRaw[headerRow];
    const rawData = xlsx.utils.sheet_to_json(sheet, { range: headerRow, defval: "", blankrows: false });
    const cleanData = rawData.map((row) => {
      const entries = Object.entries(row);
      const result = {};
      entries.forEach(([k, v], i) => {
        if(k.startsWith("__EMPTY")) {
          result[headers[i]?.toString().trim() || k] = v;
        } else {
          result[k.trim()] = v;
        }
      });
      return result;
    });

    return { data: cleanData, fileName: path.basename(filePath), filePath };
  } catch(error){
    console.error("Erro ao abrir pelo caminho:", error);
    return null;
  }
});

ipcMain.on("set-title", (event, title) => {
  BrowserWindow.getFocusedWindow().setTitle(title);
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});