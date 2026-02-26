import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const { Menu } = require('electron');
const path = require("path");
const xlsx = require("xlsx");

// Menu.setApplicationMenu(null);

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

    return {
      data: cleanData,
      fileName: path.basename(filePaths[0]),
    };
  } catch (error) {
    console.error("Erro ao ler Excel:", error);
    return null;
  }
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});