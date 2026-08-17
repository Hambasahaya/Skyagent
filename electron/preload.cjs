const { contextBridge } = require("electron");

contextBridge.exposeInMainWorld("skyagenDesktop", {
  platform: process.platform,
  version: process.versions.electron
});