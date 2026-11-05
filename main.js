const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');

// Wyłącz sprzętowe przyspieszenie GPU jeśli powoduje problemy
// app.disableHardwareAcceleration();

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    title: 'Grafik Sklepu',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
    // macOS – ukryj domyślny pasek tytułu, zostaw przyciski systemowe
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    show: false, // pokaż dopiero po załadowaniu
  });

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // Otwieraj zewnętrzne linki w domyślnej przeglądarce
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  // DevTools tylko w trybie deweloperskim
  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(() => {
  createWindow();

  // macOS: odtwórz okno po kliknięciu ikony w Docku
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ── IPC: Eksport Excel ──────────────────────────────────────────────────────
ipcMain.handle('export:excel', async (event, { storeName, year, month, summaryRows, dailyRows }) => {
  try {
    const XLSX = require('xlsx');

    // Proponowana nazwa pliku
    const pad = (n) => String(n).padStart(2, '0');
    const fileName = `Rozliczenie_${storeName}_${year}-${pad(month)}.xlsx`;

    const { filePath, canceled } = await dialog.showSaveDialog(mainWindow, {
      title: 'Zapisz plik Excel',
      defaultPath: path.join(app.getPath('documents'), fileName),
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });

    if (canceled || !filePath) return { success: false, reason: 'canceled' };

    const wb = XLSX.utils.book_new();

    // Arkusz: Rozliczenie
    const ws1 = XLSX.utils.json_to_sheet(summaryRows);
    ws1['!cols'] = Object.keys(summaryRows[0] || {}).map(k => ({ wch: Math.max(14, k.length + 2) }));
    XLSX.utils.book_append_sheet(wb, ws1, `Rozliczenie ${pad(month)}-${year}`.slice(0, 31));

    // Arkusz: Grafik dzienny
    if (dailyRows && dailyRows.length > 0) {
      const ws2 = XLSX.utils.json_to_sheet(dailyRows);
      ws2['!cols'] = [12, 22, 18, 16, 10, 10, 10, 50].map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, ws2, 'Grafik dzienny');
    }

    XLSX.writeFile(wb, filePath);
    return { success: true, filePath };
  } catch (err) {
    console.error('Export Excel error:', err);
    return { success: false, reason: err.message };
  }
});

// ── IPC: Eksport bazy JSON ──────────────────────────────────────────────────
ipcMain.handle('db:export', async (event, { data, storeName }) => {
  try {
    const fileName = `BazaDanych_${storeName}_${new Date().toISOString().slice(0,10)}.json`;
    const { filePath, canceled } = await dialog.showSaveDialog(mainWindow, {
      title: 'Eksportuj bazę danych',
      defaultPath: path.join(app.getPath('documents'), fileName),
      filters: [{ name: 'JSON', extensions: ['json'] }],
    });
    if (canceled || !filePath) return { success: false, reason: 'canceled' };
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8');
    return { success: true, filePath };
  } catch (err) {
    return { success: false, reason: err.message };
  }
});

// ── IPC: Import bazy JSON ───────────────────────────────────────────────────
ipcMain.handle('db:import', async () => {
  try {
    const { filePaths, canceled } = await dialog.showOpenDialog(mainWindow, {
      title: 'Importuj bazę danych',
      filters: [{ name: 'JSON', extensions: ['json'] }],
      properties: ['openFile'],
    });
    if (canceled || filePaths.length === 0) return { success: false, reason: 'canceled' };
    const raw = fs.readFileSync(filePaths[0], 'utf-8');
    const data = JSON.parse(raw);
    return { success: true, data };
  } catch (err) {
    return { success: false, reason: err.message };
  }
});

// ── IPC: Storage (localStorage fallback) ────────────────────────────────────
const storageFile = path.join(app.getPath('userData'), 'grafik_data.json');

ipcMain.handle('storage:get', (event, key) => {
  try {
    if (!fs.existsSync(storageFile)) return null;
    const all = JSON.parse(fs.readFileSync(storageFile, 'utf-8'));
    return all[key] ?? null;
  } catch { return null; }
});

ipcMain.handle('storage:set', (event, key, value) => {
  try {
    let all = {};
    if (fs.existsSync(storageFile)) {
      all = JSON.parse(fs.readFileSync(storageFile, 'utf-8'));
    }
    all[key] = value;
    fs.writeFileSync(storageFile, JSON.stringify(all), 'utf-8');
    return true;
  } catch { return false; }
});
