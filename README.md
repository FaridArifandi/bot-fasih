# Bot FASIH - Auto Approve

Membuka tiap link assignment FASIH di Excel, mengecek tombol di kanan bawah, dan menekan **Approve** kalau belum.

## Setup (sekali saja)
1. Jalankan `0_install_dependensi.bat`
2. Isi file `sso` (`username=...` dan `password=...`, akun SSO BPS)
3. Taruh Excel sebagai `daftar.xlsx` (kolom A = nama, kolom B = `link_fasih`; baris 1 = header)

## Menjalankan
`1_jalankan_bot.bat` lalu pilih mode:
- **[1] Dry-run**: buka link, cek status, tidak klik Approve. **Jalankan ini dulu** untuk memastikan bot mengenali tombol Approve di FASIH.
- [2] Approve 1 link, tanya y/n dulu
- [3] Batch N link
- [4] Semua antrean

Bot membuka Chrome sendiri (port 9223, profil di `chrome_debug_profile/`, terpisah dari Bot Nerwana yang memakai 9222).

## Penanganan galat
Kalau Approve tidak bisa (dialog muncul `GALAT: n`, tombol disabled, atau approve tidak terverifikasi), bot otomatis:
buka menu **CATATAN** -> aktifkan switch **Tampilkan Anomali Usaha dan Keluarga** -> Approve ulang (maks 3 percobaan per link).

## Hasil & lanjut
- Setiap link dicatat ke `hasil_approve.csv` (dan `hasil_approve.xlsx` tiap 50 link + saat selesai).
- Status `APPROVED` dan `SUDAH_APPROVE` dianggap selesai; kalau bot dijalankan lagi, link itu dilewati. Status lain dicoba ulang.
- Stop-loss: berhenti jika 3 link berturut-turut error (`GAGAL_APPROVE`, `GALAT`, `ERROR_HALAMAN`, `TIDAK_DIKENALI`). Bukti ada di `screenshots/`.
- Login SSO gagal 2x berturut-turut: bot berhenti (cek file `sso`).
