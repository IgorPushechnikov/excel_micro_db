const { app, BrowserWindow } = require('electron');
const path = require('path');

// Укажите порт, на котором запускается сервер Vite
const viteDevServerUrl = 'http://localhost:5173'; // Изменено с 3000 на 5173, и теперь это .cjs

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      // !!! ВАЖНО: nodeIntegration НЕ рекомендуется для продакшена из-за безопасности !!!
      // Для разработки с hot-reload может быть полезно, но лучше избегать.
      // preload: path.join(__dirname, 'preload.js'), // Используйте preload скрипт для безопасности
      nodeIntegration: false, // Отключаем nodeIntegration
      contextIsolation: true, // Включаем изоляцию контекста
    },
    icon: path.join(__dirname, 'public', 'icon.png'), // Укажите путь к иконке
  });

  // Режим разработки: загружаем URL с сервера Vite
  if (!app.isPackaged) { // isPackaged вернёт false в режиме разработки
    mainWindow.loadURL(viteDevServerUrl);
    // Для отладки
    mainWindow.webContents.openDevTools();
  } else {
    // Режим продакшена: загружаем локальный файл
    mainWindow.loadFile(path.join(__dirname, 'dist', 'index.html')); // Путь к файлу после сборки
  }
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// Пример: как можно вызвать команду из renderer процесса (React App)
// через IPC (Inter-Process Communication)
// const { ipcMain } = require('electron');
// 
// ipcMain.handle('ping', (event, message) => {
//   console.log('Ping received:', message);
//   return 'Pong';
// });
