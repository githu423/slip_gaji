# Generator Slip Gaji

Website statis (tanpa framework, tanpa backend) untuk mengubah satu file Excel
data gaji banyak pegawai menjadi surat rincian gaji (PDF) per orang — bisa
diunduh satu-satu atau sekaligus dalam satu file ZIP.

Semua proses (baca Excel, generate PDF, bikin ZIP) berjalan di browser
pengguna lewat JavaScript. Tidak ada data yang dikirim ke server mana pun,
jadi cukup di-host sebagai file statis.

## Isi folder
- `index.html` — struktur halaman dan konten
- `css/styles.css` — seluruh tampilan, komponen, dan layout responsif
- `js/app.js` — unggah Excel, pencarian, perhitungan gaji, dan ekspor dokumen
- `vercel.json` — konfigurasi minimal untuk Vercel
- `README.md` — file ini

## Cara pakai (di website)
1. Buka `index.html`, klik area upload dan pilih file Excel (`.xlsx`).
   Kalau belum punya, klik "Unduh contoh format Excel" untuk lihat contohnya.
2. Kolom yang dikenali otomatis:
   - `Nama`, `NIK`, `Jabatan`, `Departemen` → info pegawai
   - `Gaji Pokok` → komponen gaji pokok
   - kolom apa pun yang diawali kata **Tunjangan** → otomatis ditambahkan
   - kolom apa pun yang diawali kata **Potongan** → otomatis dikurangkan
3. Isi nama perusahaan, alamat, periode gaji, dan pilih **format surat** (Word
   atau PDF) di langkah 2.
4. Di tabel langkah 3, cek datanya (kolom **Gaji Kotor** dan **Gaji Bersih**
   sudah otomatis dihitung), lalu:
   - klik **Unduh** di satu baris untuk slip satu pegawai saja, atau
   - klik tombol **Unduh semua**:
     - format **Word** → menghasilkan **satu file `.docx`**, tiap pegawai
       otomatis mulai di halaman baru (1 lembar = 1 orang), tinggal dicetak
       atau diedit lagi di Microsoft Word,
     - format **PDF** → menghasilkan satu file **ZIP** berisi PDF terpisah
       per pegawai.

Setiap slip menampilkan: data pegawai, rincian **Pendapatan** (Gaji Pokok +
semua kolom Tunjangan) dengan totalnya sebagai **Gaji Kotor**, rincian
**Potongan** (jika ada), dan **Gaji Bersih** di baris paling bawah.

## Cara deploy ke Vercel

**Opsi 1 — lewat Vercel CLI**
```bash
npm i -g vercel
cd slip-gaji
vercel
```
Ikuti instruksi di terminal (login, pilih scope, deploy). Vercel akan
mendeteksi ini sebagai proyek statis, tidak perlu build command apa pun.

**Opsi 2 — lewat dashboard Vercel**
1. Push folder ini ke repo GitHub (sertakan folder `css` dan `js`).
2. Di [vercel.com](https://vercel.com), klik **Add New → Project**, pilih repo tersebut.
3. Framework Preset: pilih **Other**. Build Command & Output Directory kosongkan saja.
4. Klik **Deploy**.

## Catatan
- Library yang dipakai (SheetJS, jsPDF, JSZip, docx) dimuat lewat CDN —
  pastikan koneksi internet aktif saat halaman dibuka.
- Library Word (`docx`) dikunci ke versi 8.5.0 supaya berkas UMD-nya stabil
  di CDN; tidak perlu diubah kecuali memang ingin upgrade.
- Kalau butuh kolom tunjangan/potongan yang jumlahnya beda-beda tiap
  perusahaan, tinggal tambah kolom baru di Excel dengan awalan "Tunjangan"
  atau "Potongan" — tidak perlu ubah kode.

## Pengembangan lokal

Tidak memerlukan build atau instalasi dependensi. Jalankan server statis dari root proyek:

```bash
python3 -m http.server 8000 --bind 0.0.0.0
```

Buka `http://localhost:8000`. Ubah tampilan di `css/styles.css` (dikelompokkan menjadi token desain, header, kartu, form, tabel, footer, dan responsivitas). Logika aplikasi ada di `js/app.js`.

### Pemeriksaan manual
- Unduh template, lalu unggah kembali; kedua pegawai dan total gajinya harus tampil.
- Ketik beberapa karakter pada pencarian; fokus dan posisi kursor harus tetap terjaga.
- Coba file kosong, kolom wajib yang hilang, dan file bukan Excel; pesan kesalahan harus muncul.
- Isi detail slip, unduh Word/PDF per pegawai dan semua pegawai, lalu periksa hasilnya.
- Cek pada layar ponsel dan navigasikan tombol unggah dengan Tab, Enter, dan Spasi.
