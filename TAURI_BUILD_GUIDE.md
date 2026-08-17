# SKYagen Desktop Build Guide

Project frontend sudah disiapkan sebagai aplikasi desktop Tauri v2.

## Script

```powershell
npm run tauri:dev
npm run tauri:build
```

Output build Windows akan berada di:

```text
src-tauri/target/release/bundle/
```

## Requirement Windows

Install dependency Tauri berikut sebelum build `.exe`:

1. Rust toolchain: https://rustup.rs
2. Microsoft C++ Build Tools atau Visual Studio Build Tools dengan workload C++.
3. WebView2 Runtime, biasanya sudah ada di Windows 10/11 modern.

Setelah install Rust, tutup dan buka terminal baru, lalu cek:

```powershell
cargo --version
rustc --version
```

Jika sudah muncul versi, jalankan:

```powershell
cd C:\Users\ASUS\Documents\js\Skyagents
npm install
npm run tauri:build
```

## Requirement macOS

Build app macOS harus dilakukan dari mesin macOS.

```bash
xcode-select --install
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
cd /path/to/Skyagents
npm install
npm run tauri:build
```

Output macOS akan berada di folder bundle Tauri, biasanya `.app` dan `.dmg` jika target bundle tersedia.

## Backend

Aplikasi desktop tetap memakai API backend SKYagen. Default FE mengarah ke:

```text
http://localhost:8080
```

Jika backend tidak berada di localhost, gunakan popup Base URL yang sudah ada di app. Nilai akan disimpan di localStorage Tauri WebView.