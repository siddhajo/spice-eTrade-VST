// Electron main process — wraps the existing Express server in a
// desktop window. The architecture: the Node main process boots
// server.js in-process (same event loop, simplest possible model),
// then a BrowserWindow points at http://localhost:<PORT>/.
//
// Why in-process and not spawn(): server.js is pure-JS (no native
// deps), shares a single sql.js memory image with this process, and
// avoids the Windows multi-process firewall / antivirus prompt the
// user would otherwise see on first launch.

const { app, BrowserWindow, Menu, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const net = require('net');

const PORT = Number(process.env.PORT) || 3001;

// ─── 1. Per-user writable data directory ─────────────────────────
// Set BEFORE requiring server.js so db.js + multer pick up the right
// path. `app.getPath('userData')` resolves to:
//   Windows: %APPDATA%\Spice e-Trade
//   macOS:   ~/Library/Application Support/Spice e-Trade
//   Linux:   ~/.config/Spice e-Trade
const userData = app.getPath('userData');
const dataDir  = path.join(userData, 'data');
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
process.env.SPICE_DATA_DIR = dataDir;
process.env.PORT = String(PORT);

// ─── 2. Boot the existing Express server ─────────────────────────
// Wrap in try/catch so any startup error surfaces in a dialog
// instead of vanishing into the macOS console.
try {
  require(path.join(__dirname, '..', 'server.js'));
} catch (e) {
  app.whenReady().then(() => {
    dialog.showErrorBox('Startup failed',
      'The server module failed to load.\n\n' + (e.stack || e.message));
    app.quit();
  });
}

// ─── 3. Wait for the server to accept connections ────────────────
// The server boots asynchronously (sql.js loads its WASM via fetch);
// poll the TCP port instead of guessing a delay.
function waitForServer(port, attempts = 80) {
  return new Promise((resolve, reject) => {
    const tick = () => {
      const s = net.createConnection({ port }, () => { s.end(); resolve(); });
      s.on('error', () => {
        if (--attempts <= 0) return reject(new Error(`Server failed to start on port ${port}`));
        setTimeout(tick, 250);
      });
    };
    tick();
  });
}

// ─── 4. Window state persistence ─────────────────────────────────
// Saves last position + size to userData so the next launch opens
// the window where the user left it. No external dep — small JSON.
const stateFile = path.join(userData, 'window-state.json');
function loadWindowState() {
  try { return JSON.parse(fs.readFileSync(stateFile, 'utf8')); } catch { return null; }
}
function saveWindowState(win) {
  if (!win || win.isDestroyed()) return;
  const b = win.getBounds();
  try { fs.writeFileSync(stateFile, JSON.stringify({ ...b, maximized: win.isMaximized() })); } catch {}
}

let win = null;
async function createWindow() {
  const saved = loadWindowState() || {};
  win = new BrowserWindow({
    width:  saved.width  || 1400,
    height: saved.height || 900,
    x:      saved.x,  // undefined = center on primary display
    y:      saved.y,
    minWidth: 1100,
    minHeight: 700,
    title: 'Spice e-Trade',
    show: false,             // wait until URL is loaded to avoid white flash
    backgroundColor: '#F5F1EB',
    icon: path.join(__dirname, 'icon.png'),  // optional; ignored if missing
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  if (saved.maximized) win.maximize();

  try { await waitForServer(PORT); }
  catch (err) {
    dialog.showErrorBox('Startup failed', err.message);
    app.quit();
    return;
  }

  win.loadURL(`http://localhost:${PORT}/`);
  win.once('ready-to-show', () => win.show());

  // Persist size/position on every relevant event.
  win.on('close', () => saveWindowState(win));
  win.on('resize', () => saveWindowState(win));
  win.on('move',   () => saveWindowState(win));

  // External http(s) links open in the OS default browser; everything
  // local (the receipt-print popups your app opens via window.open)
  // gets a real Electron BrowserWindow so the native print dialog
  // works.
  win.webContents.setWindowOpenHandler(({ url }) => {
    const isLocal = url.startsWith(`http://localhost:${PORT}`) ||
                    url.startsWith('about:blank');
    if (!isLocal) {
      shell.openExternal(url);
      return { action: 'deny' };
    }
    return { action: 'allow' };
  });

  // Open download dialogs for any HTTP response with Content-Disposition
  // attachment (PDF/XLSX/XML downloads). Electron handles this natively;
  // no extra wiring needed here — left as a hook in case we want to
  // route saves to a default folder later.
}

// ─── 5. Native menu (mac-friendly + Windows fallback) ────────────
function buildMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac ? [{
      label: app.name,
      submenu: [
        { role: 'about' }, { type: 'separator' },
        { role: 'services' }, { type: 'separator' },
        { role: 'hide' }, { role: 'hideOthers' }, { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' },
      ],
    }] : []),
    {
      label: 'File',
      submenu: [
        {
          label: 'Open Data Folder',
          click: () => shell.openPath(dataDir),
        },
        {
          label: 'Open Backups Folder',
          click: () => shell.openPath(path.join(dataDir, 'backups')),
        },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' }, { role: 'redo' }, { type: 'separator' },
        { role: 'cut' }, { role: 'copy' }, { role: 'paste' },
        { role: 'selectAll' },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' }, { role: 'forceReload' }, { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
    {
      label: 'Window',
      submenu: [
        { role: 'minimize' }, { role: 'zoom' },
        ...(isMac ? [{ type: 'separator' }, { role: 'front' }] : [{ role: 'close' }]),
      ],
    },
    {
      role: 'help',
      submenu: [
        {
          label: 'About Spice e-Trade',
          click: () => dialog.showMessageBox({
            type: 'info',
            title: 'About',
            message: `Spice e-Trade v${app.getVersion()}`,
            detail: `Data folder:\n${dataDir}`,
          }),
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ─── 6. Single-instance lock ─────────────────────────────────────
// Prevents two copies of the server fighting over port 3001 + the
// same SQLite file. A second double-click brings the existing window
// to the front instead.
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (win) {
      if (win.isMinimized()) win.restore();
      win.focus();
    }
  });

  app.whenReady().then(() => {
    buildMenu();
    createWindow();
  });

  // ─── 7. Lifecycle ─────────────────────────────────────────────
  app.on('window-all-closed', () => {
    // macOS apps usually stay running with no windows. On other
    // platforms, quitting on window-close is the expected UX.
    if (process.platform !== 'darwin') app.quit();
  });

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });

  // Best-effort flush of pending DB writes before exit. db.js already
  // hooks SIGINT/SIGTERM/beforeExit, but Electron's `before-quit`
  // fires earlier than process exit signals on some platforms.
  app.on('before-quit', () => {
    try { require('../db').flushSave(); } catch {}
  });
}
