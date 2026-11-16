# SPK Prioritas Monitoring Progres LINTOR — Flask

Prototipe akademik untuk membantu memeriksa progres LINTOR, membandingkan desa, dan menyusun **indikasi prioritas monitoring** dari data snapshot yang tersedia. Tampilan dibuat dengan identitas visual merah-putih dan nuansa administratif; aplikasi ini bukan aplikasi resmi ATR/BPN.

## Login dan logout
- Saat membuka aplikasi, pengguna akan diarahkan ke halaman login.
- Kredensial demo default: username `admin`, password `Lintor@2026`.
- Klik tombol keluar di bagian bawah sidebar untuk logout.
- Login default hanya untuk demo lokal. Ganti sebelum digunakan bersama:
  - Windows CMD: `set LINTOR_USERNAME=nama_pengguna`
  - Windows CMD: `set LINTOR_PASSWORD=kata_sandi_yang_kuat`
  - Opsional, tetapkan secret sesi: `set FLASK_SECRET_KEY=string_acak_panjang`
- Variabel di atas perlu ditetapkan pada terminal yang sama sebelum menjalankan `python app.py`. Jangan menyebarkan kredensial demo atau menjalankan aplikasi di jaringan publik tanpa pengamanan tambahan.

## Isi proyek
- `app.py`: aplikasi Flask, pembacaan Excel, indikator, metode SAW, ekspor CSV.
- `data/progres_lintor.xlsx`: salinan file sumber yang diberikan.
- `templates/` dan `static/`: halaman web responsif.
- `run_windows.bat`: cara cepat menjalankan pada Windows.

## Cara menjalankan (Windows)
1. Ekstrak ZIP ke folder yang mudah ditemukan.
2. Pastikan Python 3.10 atau lebih baru terpasang. Saat memasang Python, aktifkan opsi **Add Python to PATH**.
3. Klik dua kali `run_windows.bat`, atau jalankan perintah berikut dari folder proyek:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

4. Buka `http://127.0.0.1:5000` di browser.
5. Untuk berhenti, tekan `Ctrl+C` di terminal.

## Metode dan asumsi
Versi awal memakai SAW dengan bobot demonstrasi 40% (rasio K1 rendah), 35% (sisa target), dan 25% (selisih Pengumuman–K1). Bobot dan kriteria **belum disahkan oleh kantor**. Skor hanya alat bantu untuk memeriksa data, bukan keputusan resmi atau penilaian kinerja petugas.

- Rasio K1 dihitung sebagai `K1 / Target × 100`, jika target lebih dari nol.
- Sisa target dihitung sebagai `max(Target − K1, 0)`.
- Selisih Pengumuman–K1 adalah indikator diagnostik awal, bukan definisi hambatan resmi.
- Kriteria yang saling berkaitan dapat membuat suatu isu terhitung lebih dari sekali. Validasi dan penyepakatan kriteria perlu dilakukan dengan pembimbing/petugas.
- Periode data dibaca otomatis dari judul file Excel yang diimpor (contoh: `(21 OKTOBER 2026)`). Jika tanggal tidak ditemukan, dipakai waktu terakhir file diimpor. Aplikasi tidak mengarang tren bulanan atau prediksi akhir tahun karena data historis belum tersedia.
- Nilai kosong tidak otomatis disamakan dengan nol. Nilai nol perlu dikonfirmasi artinya kepada pemilik data.

## Visualisasi dashboard
- Dashboard menggunakan satu area visualisasi yang dapat diganti antara **Perbandingan tahapan** (diagram batang/garis target) dan **Komposisi tahapan** (pie chart), sehingga tidak menampilkan diagram serupa secara bersamaan.
- Filter program, kecamatan, dan desa diterapkan pada ringkasan, diagram, prioritas, tabel data, serta ekspor.
- Diagram per desa dibatasi maksimal 30 baris agar tetap mudah dibaca dan lebih ringan. Halaman Data Progres tetap menyediakan daftar data lengkap.
- Chart.js disertakan lokal (`static/js/chart.umd.js`), sehingga diagram tetap tampil tanpa internet.

## Format Excel
Aplikasi mencoba mencari baris header secara otomatis. Kolom wajib: `Program`, `Desa`, `Kecamatan`, `Target`, `Pemberkasan`, `SU`, `Pengumuman`, `Pengesahan Pengumuman`, `K1`. Baris `Jumlah`/`Total` diabaikan karena total dihitung ulang dari baris desa untuk menghindari penghitungan ganda.

## Batasan keamanan
Login/logout dasar tersedia, tetapi prototipe ini belum memiliki manajemen banyak pengguna, pembatasan percobaan login, audit log, backup otomatis, perlindungan CSRF menyeluruh, atau hardening produksi. Jangan mengunggah data pribadi/rahasia atau menaruh aplikasi di jaringan publik sebelum mendapat izin dan menambahkan kontrol keamanan yang sesuai.

### Pembaruan filter data

Filter Program, Kecamatan, dan Desa/Kelurahan mengikuti baris valid pada file Excel yang terakhir diunggah. Pilihan Program akan memperbarui pilihan Kecamatan dan Desa; pilihan Kecamatan akan memperbarui pilihan Desa. Setiap perubahan pilihan langsung menerapkan filter. Jika pilihan lokasi lama tidak sesuai setelah Program/Kecamatan berubah, pilihan turunannya otomatis kembali ke “Semua” agar hasil tidak keliru. Statistik, visualisasi, daftar prioritas, data progres, dan ekspor menggunakan filter yang sama.


## Deploy gratis ke Render (disarankan untuk demo)

Proyek sudah disiapkan untuk dijalankan dengan Gunicorn melalui `Procfile`.
File `requirements.txt` juga sudah memuat Gunicorn dan `.python-version` menetapkan Python 3.12.10.

### Persiapan sebelum deploy
1. Buat akun GitHub dan repository baru, misalnya `spk-lintor`.
2. Ekstrak ZIP ini. Unggah **isi folder `spk_lintor`** ke repository (pastikan `app.py`, `requirements.txt`, `Procfile`, folder `templates`, `static`, dan `data` berada di root repository, bukan di dalam folder bersarang).
3. Pastikan `data/progres_lintor.xlsx` ikut terunggah.
4. Jangan unggah file `.venv`, `__pycache__`, atau kredensial pribadi.

### Membuat layanan di Render
1. Buka https://render.com/ lalu daftar/login.
2. Hubungkan akun GitHub ke Render.
3. Pilih **New + → Web Service**.
4. Pilih repository `spk-lintor` yang sudah dibuat.
5. Isi konfigurasi:
   - **Name:** `spk-lintor` (atau nama lain yang tersedia)
   - **Language/Runtime:** Python
   - **Branch:** `main` (atau branch yang dipakai)
   - **Root Directory:** kosongkan jika `app.py` berada di root repository
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --workers 1 --threads 4 --timeout 120`
   - **Instance Type:** Free
6. Pada bagian **Environment Variables**, tambahkan:
   - `FLASK_SECRET_KEY` = string acak panjang (minimal 32 karakter; jangan gunakan contoh publik)
   - `LINTOR_USERNAME` = nama pengguna yang bukan `admin`
   - `LINTOR_PASSWORD` = kata sandi kuat yang berbeda dari kata sandi lain
   - `COOKIE_SECURE` = `1`
7. Klik **Create Web Service** dan tunggu proses build/deploy selesai.
8. Jika status sudah **Live**, buka URL `https://nama-layanan.onrender.com`.
9. Masuk memakai nilai `LINTOR_USERNAME` dan `LINTOR_PASSWORD` yang Anda atur di Environment Variables.

### Membuat secret key
Jalankan perintah berikut di komputer sendiri, lalu salin hasilnya ke `FLASK_SECRET_KEY` di Render:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

### Batasan penting paket gratis Render
- Layanan gratis akan tidur setelah sekitar 15 menit tanpa trafik. Kunjungan pertama setelah tidur dapat memerlukan waktu sekitar satu menit untuk aktif kembali.
- Sistem berkas layanan gratis bersifat sementara. File Excel yang diunggah melalui menu **Impor Data** dapat hilang ketika layanan tidur, restart, atau deploy ulang. File Excel bawaan di repository akan tersedia kembali setelah aplikasi dimulai ulang.
- Karena itu, paket gratis ini cocok untuk demo/presentasi dan uji coba, **bukan** penyimpanan data operasional. Jika data Excel harus selalu tersimpan setelah diunggah, perlu penyimpanan eksternal persisten atau layanan berbayar; jangan mengandalkan folder `data/` pada Render Free.
- Paket gratis bukan jaminan layanan selalu aktif dan bukan pilihan untuk data rahasia/operasional kantor. Gunakan data yang telah mendapat izin untuk dipublikasikan.

### Jika deploy gagal
- Buka tab **Logs** pada layanan Render.
- `ModuleNotFoundError`: pastikan `requirements.txt` berada di root repository dan Build Command benar.
- `Application failed to respond`: periksa Start Command harus `gunicorn app:app --workers 1 --threads 4 --timeout 120`.
- Data tidak terbaca: pastikan `data/progres_lintor.xlsx` ikut diunggah dan nama/path-nya sama.
- Login gagal: periksa kembali `LINTOR_USERNAME` dan `LINTOR_PASSWORD` di Environment Variables, lalu lakukan **Manual Deploy → Deploy latest commit** setelah perubahan kode.

Dokumentasi resmi:
- https://render.com/docs/deploy-flask
- https://render.com/docs/free
