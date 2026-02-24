const { app, BrowserWindow, Menu, shell, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 800,
    minHeight: 600,
    title: 'Rupaiya - Personal Expense Tracker',
    icon: path.join(__dirname, 'assets', process.platform === 'win32' ? 'icon.ico' : 'icon.png'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    backgroundColor: '#F4F5F8',
    show: false,
  });

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    // Open DevTools in dev mode
    if (process.argv.includes('--dev')) {
      mainWindow.webContents.openDevTools();
    }
  });

  mainWindow.on('closed', () => { mainWindow = null; });
}

/* ============================================================ APP MENU */
function buildMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        { label: 'New Expense\tN', accelerator: 'N', click: () => mainWindow?.webContents.executeJavaScript("navigate('add')") },
        { type: 'separator' },
        { label: 'Import Excel...', accelerator: 'CmdOrCtrl+I', click: () => mainWindow?.webContents.executeJavaScript("navigate('import')") },
        { label: 'Export to Excel', accelerator: 'CmdOrCtrl+E', click: () => mainWindow?.webContents.executeJavaScript("exportAllExcel()") },
        { type: 'separator' },
        { role: 'quit', label: 'Exit' }
      ]
    },
    {
      label: 'View',
      submenu: [
        { label: 'Dashboard', accelerator: 'CmdOrCtrl+1', click: () => mainWindow?.webContents.executeJavaScript("navigate('dashboard')") },
        { label: 'Expenses', accelerator: 'CmdOrCtrl+2', click: () => mainWindow?.webContents.executeJavaScript("navigate('expenses')") },
        { label: 'Analytics', accelerator: 'CmdOrCtrl+3', click: () => mainWindow?.webContents.executeJavaScript("navigate('analytics')") },
        { label: 'Custom Charts', accelerator: 'CmdOrCtrl+4', click: () => mainWindow?.webContents.executeJavaScript("navigate('custom-charts')") },
        { type: 'separator' },
        { role: 'reload' },
        { role: 'togglefullscreen' },
        { role: 'toggleDevTools', label: 'Toggle DevTools' },
      ]
    },
    {
      label: 'Window',
      submenu: [
        { role: 'minimize' },
        { role: 'zoom' },
      ]
    },
    {
      label: 'Help',
      submenu: [
        {
          label: 'About Rupaiya',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: 'About Rupaiya',
              message: 'Rupaiya - Personal Expense Tracker',
              detail: 'Version 1.0.0\nFree Forever\n\nAll data stored locally on your device.\nNo internet connection required.',
            });
          }
        }
      ]
    }
  ];

  if (process.platform === 'darwin') {
    template.unshift({ label: app.name, submenu: [{ role: 'about' }, { type: 'separator' }, { role: 'quit' }] });
  }

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/* ============================================================ LIFECYCLE */
app.whenReady().then(() => {
  createWindow();
  buildMenu();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
