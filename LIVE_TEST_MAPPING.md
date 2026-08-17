# SKYagen Live Test Mapping

Tujuan file ini adalah membuat live test tetap cepat, terarah, dan tanpa mock. Jalankan fast path dulu; lanjut ke CRUD/import/export hanya jika fast path sudah hijau.

## Environment

- Backend default: `http://localhost:8080`
- Frontend dev: `http://127.0.0.1:5173`
- Override FE base URL disimpan di localStorage: `skyagen_api_base_url`
- Status disconnected disimpan di localStorage: `skyagen_api_disconnected_at` dan `skyagen_api_disconnected_page`
- Token FE disimpan di localStorage: `skyagen_access_token`

## Fast Path, 5-8 Menit

| Urutan | Area | FE View | Endpoint utama | Kriteria lolos |
| --- | --- | --- | --- | --- |
| 1 | Koneksi API | App startup / popup API | `GET /api/dashboard/summary` atau request awal halaman | Jika backend mati, popup base URL muncul dan URL tersimpan |
| 2 | Auth | Login | `POST /api/auth/login`, `GET /api/auth/me` | Login admin berhasil, token tersimpan |
| 3 | Dashboard | Dashboard | `GET /api/dashboard/summary`, `/charts`, `/recent-documents`, `/recent-joinings` | Semua widget tampil; data kosong tampil elegan |
| 4 | Seafarers | Crew database/search | `GET /api/seafarers`, `GET /api/seafarers/search` | List/search jalan tanpa fallback mock |
| 5 | Principal & Vessel | Principal/Vessel page | `GET /api/admin/principals`, `GET /api/admin/vessels` | Relasi data bisa dibaca |
| 6 | Joining | Joining principal page | `GET /api/joining-principals`, `GET /api/admin/joining-statuses` | List/status tampil |
| 7 | Documents | Document pages | `GET /api/admin/document-types`, `/document-names`, `/documents/report`, `/documents/expiring` | Master dan report terbaca |
| 8 | Reports | Reports page | `GET /api/reports/crew`, `/documents`, `/joining` | Summary report terbaca |
| 9 | Admin | Users/settings | `GET /api/admin/users`, `GET /api/settings/profile`, `GET /api/admin/settings/app` | Admin data terbaca |

## Mapping Per Modul

| Modul FE | Endpoint baca cepat | Endpoint CRUD/aksi yang dites setelah fast path |
| --- | --- | --- |
| Auth | `POST /api/auth/login`, `GET /api/auth/me` | `POST /api/auth/register`, `PUT /api/auth/password` |
| Dashboard | `GET /api/dashboard/summary`, `/charts`, `/recent-documents`, `/recent-joinings` | - |
| Seafarers | `GET /api/seafarers`, `GET /api/seafarers/search`, `GET /api/seafarers/:id` | `POST /api/seafarers`, `PUT /api/seafarers/:id`, `DELETE /api/seafarers/:id` |
| Emergency contacts | ikut payload seafarers | ikut CRUD seafarers, tidak ada API terpisah |
| Seafarer documents | ikut detail/payload seafarers | ikut CRUD seafarers dan master document |
| Document masters | `GET /api/admin/document-types`, `GET /api/admin/document-names` | CRUD document types/names |
| Principals | `GET /api/admin/principals`, `GET /api/admin/principals/:id/detail` | CRUD principals, custom fields, requirements |
| Vessels | `GET /api/admin/vessels` | CRUD vessels dan relasi principal |
| Joining principals | `GET /api/joining-principals`, search/filter | CRUD joining, status, history |
| Requirements | `GET /api/admin/principal-requirements`, `GET /api/principal-requirements/check` | CRUD requirements |
| Blacklist | `GET /api/admin/blacklists` | create/release/delete blacklist |
| Operations | - | `POST /api/operations/sign-on`, `POST /api/operations/sign-off` |
| Import | - | `POST /api/admin/import/seafarers`, `/principals`, `/vessels`, `/joining-principals` |
| Export | - | `GET /api/admin/export/seafarers`, `/principals`, `/vessels`, `/joining-principals` |
| Reports | `GET /api/reports/crew`, `/documents`, `/joining` | `GET /api/reports/export?type=crew|documents|joining` |
| Users | `GET /api/admin/users` | CRUD users, status, password |
| Settings | `GET /api/settings/profile`, `GET /api/admin/settings/app` | update profile/app settings |

## Yang Ditunda Saat Fast Path

- Import Excel besar.
- Export semua kolom untuk semua modul sekaligus.
- Delete permanen data produksi.
- CRUD custom field lengkap satu per satu.
- Requirement matrix semua principal/vessel jika data sudah besar.

## Command Smoke Test Cepat

```powershell
cd C:\Users\ASUS\Documents\js\Skyagents
$env:API_BASE_URL='http://localhost:8080'
$env:ADMIN_EMAIL='admin@example.com'
$env:ADMIN_PASSWORD='password'
npm.cmd run live:smoke
```

Script ini tidak memakai mock dan tidak membuat/menghapus data. Ia hanya login lalu membaca endpoint kritis dengan timeout pendek.
