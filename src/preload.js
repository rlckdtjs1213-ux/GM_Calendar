'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  getConfig: () => ipcRenderer.invoke('config:get'),
  pickFolder: () => ipcRenderer.invoke('config:pickFolder'),
  setTeam: (team) => ipcRenderer.invoke('team:set', team),
  reload: () => ipcRenderer.invoke('events:reload'),
  onEventsUpdate: (callback) =>
    ipcRenderer.on('events:update', (_e, payload) => callback(payload)),
});
