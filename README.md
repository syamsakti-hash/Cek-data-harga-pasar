# Perbandingan Harga Pasar Kendaraan

Aplikasi PHP sederhana untuk mengunggah dua file harga kendaraan, membandingkan data berdasarkan `VEHICLE_ID`, melihat pratinjau hasil, dan mengunduh laporan Excel `.xlsx`.

## Fitur

- Menerima File A (existing) dan File B (eksternal) dalam format CSV, XLSX, atau XLS.
- Mendukung CSV dengan delimiter titik koma (`;`) atau koma (`,`).
- Mencocokkan kendaraan dengan `VEHICLE_ID`.
- Menghitung `harga_baru` sebagai rata-rata harga A dan B ketika ID ada pada keduanya.
- Menghitung `selisih_perubahan` sebagai harga File B dikurangi harga File A. Nilai positif berarti naik; nilai negatif berarti turun.
- Memasukkan ID yang hanya ada di File B, menggunakan harga File B sebagai `harga_baru`.
- Memasukkan ID yang hanya ada di File A, menggunakan harga File A sebagai `harga_baru`.
- Menampilkan pratinjau dan mengunduh workbook Excel.

## Format file

File A memerlukan kolom `VEHICLE_ID`, `VEHICLE_MODEL`, `TAHUN`, dan `HARGA KENDARAAN SEHATI`.

File B memerlukan kolom `VEHICLE_ID`, `VEHICLE_MODEL`, `TAHUN`, dan `HARGA PASAR OLX`.

CSV dengan header seperti contoh pengguna (dipisahkan titik koma) didukung. Excel hasil mengikuti template `TEMPLATE HASIL.csv`: `vehicle_id`, `vehicle_model`, `tahun`, `harga_kendaraan_sehati`, `harga_pasar_olx`, `harga_baru`, `selisih_perubahan`, `status_perubahan`, `STATUS`.

## Persyaratan

- PHP 8.1+
- Composer
- Ekstensi PHP `mbstring`, `xml`, `zip`, `gd` (umumnya dibutuhkan PhpSpreadsheet; pastikan ekstensi yang diminta Composer tersedia)

## Menjalankan lokal

```bash
composer install
php -S localhost:8000
```

Buka http://localhost:8000, pilih File A dan File B, lalu klik **Bandingkan Harga**. Setelah tabel pratinjau tampil, klik **Unduh Hasil Excel (.xlsx)**.

## Catatan perhitungan

- Untuk data yang ada di kedua file, `harga_baru = (harga_kendaraan_sehati + harga_pasar_olx) / 2`.
- `selisih_perubahan = harga_pasar_olx - harga_kendaraan_sehati`.
- Untuk kendaraan hanya di File B, `harga_baru` mengambil `harga_pasar_olx`, selisih tidak tersedia dan status perubahan adalah `Data baru`.
- Untuk kendaraan hanya di File A, `harga_baru` mengambil `harga_kendaraan_sehati`, selisih tidak tersedia dan status perubahan adalah `Tidak ada pembanding`.
