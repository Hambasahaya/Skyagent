# SKYagen Electron Build Guide

Project sudah disiapkan untuk Electron tanpa menghapus konfigurasi Tauri.

## Script

```powershell
npm run electron:dev
npm run electron:pack
npm run electron:build
```

- `electron:dev`: build Vite lalu membuka app desktop.
- `electron:pack`: membuat folder app unpacked, lebih cepat untuk validasi.
- `electron:build`: membuat installer Windows `.exe` dengan electron-builder.

Output Windows:

```text
release/
```

## Backend

Electron tetap memakai backend SKYagen melalui API. Default FE mengarah ke:

```text
http://localhost:8080
```

Jika backend berbeda, gunakan popup Base URL di aplikasi.

## Jika Electron Download Lambat atau Stuck

Electron perlu mendownload binary Chromium/Electron ke `node_modules/electron/dist`.
Jika `npx electron --version` atau `npm run electron:pack` menggantung, jalankan ulang install dengan mirror:

```powershell
$env:ELECTRON_MIRROR="https://npmmirror.com/mirrors/electron/"
npm install electron --save-dev
npx electron --version
```

Jika sudah muncul versi Electron, lanjut:

```powershell
npm run electron:pack
npm run electron:build
```

## macOS

Build `.dmg` harus dilakukan di macOS:

```bash
npm install
npm run electron:build
```