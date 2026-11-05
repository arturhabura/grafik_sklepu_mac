const { contextBridge, ipcRenderer } = require('electron');

// Bezpieczny most między procesem głównym a renderującym
contextBridge.exposeInMainWorld('api', {
  export: {
    excel: (args) => ipcRenderer.invoke('export:excel', args),
  },
  db: {
    export: (args) => ipcRenderer.invoke('db:export', args),
    import: ()    => ipcRenderer.invoke('db:import'),
  },
  storage: {
    get: (key)        => ipcRenderer.invoke('storage:get', key),
    set: (key, value) => ipcRenderer.invoke('storage:set', key, value),
  },
  isElectron: true,
});
