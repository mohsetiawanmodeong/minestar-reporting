# Excel Data Cleaner untuk Operasi Penambangan MINESTAR

Aplikasi web ini memproses file Excel data operasi penambangan mentah dan membersihkannya sesuai dengan persyaratan format khusus, serta membuat laporan Performance Summary.

## Fitur

- Mengkonversi data mesin ke format standar
- Memformat ulang informasi tanggal/waktu ke MM/DD/YYYY H:MM:SS untuk kompatibilitas filter Excel
- Membuat tabel Excel dengan filter otomatis dan formatting yang profesional
- Mengkonversi durasi dari detik ke jam (2 angka desimal)
- Mengkategorikan jenis delay berdasarkan awalan:
  - D- → DELAY
  - S- → STANDBY
  - UX- → UNPLANNED DOWN
  - X- → PLANNED DOWN
  - XX- → EXTENDED LOSS
- Membuat sheet Performance Summary otomatis yang menghitung:
  - Operating Hours per unit
  - Total waktu per kategori delay
  - Performance Availability (PA) menggunakan rumus khusus
- Dukungan bilingual (Bahasa Indonesia dan Inggris)
- Antarmuka pengguna yang responsif dengan indikator loading
- Styling tabel Excel otomatis dengan warna per jenis sheet:
  - Orange untuk data Delay
  - Kuning/orange muda untuk data Cycle
  - Biru untuk data Performance

## Instalasi

1. Kloning atau unduh repositori ini
2. Instal dependensi yang diperlukan:

```bash
pip install -r requirements.txt
```

## Penggunaan

1. Mulai aplikasi:

```bash
python app.py
```

2. Buka browser web dan kunjungi `http://127.0.0.1:5000`
3. Ikuti petunjuk proses data:
   - Filter dan pilih hanya data LHD Production (tanpa RB, 730, dan LHD Development)
   - Tarik data untuk 1 hari penuh (24 jam)
   - Atur rentang waktu dengan Start Time: 00:00 dan Finish Time: 23:59
   - Copy data Cycle dan Delay dari Minestar dan paste ke file template Excel
   - Proses secara terpisah untuk setiap site (GBC & DMLZ)
4. Unggah file Excel yang telah disiapkan
5. Aplikasi akan memproses file dan mengembalikan versi yang telah dibersihkan untuk diunduh

## Format Input

Aplikasi ini mengharapkan file Excel dengan kolom-kolom berikut:
- Sheet Delay:
  - Machine/Unit 
  - Start Date (dalam format seperti "Fri Mar 28 07:25:25 WIT 2025")
  - Finish Date (dalam format seperti "Fri Mar 28 10:04:30 WIT 2025")
  - Duration (dalam format seperti "9,545 s")
  - Delay Type (dengan awalan seperti "D-", "S-", dll.)
  - Description

- Sheet Cycle:
  - Unit
  - Operator
  - Start time (format serupa dengan Start Date pada Delay)
  - Finish Time (format serupa dengan Finish Date pada Delay)
  - Dur
  - Source
  - Destination

## Format Output

File Excel yang diproses akan memiliki:

### Sheet Delay
- Unit (dari Machine)
- Start (diformat ulang sebagai MM/DD/YYYY H:MM:SS)
- Finish (diformat ulang sebagai MM/DD/YYYY H:MM:SS)
- Dur (durasi dalam jam, 2 angka desimal)
- Desc (deskripsi)
- Delay Type (dipertahankan dari input)
- Category (ditentukan berdasarkan awalan Delay Type)

### Sheet Cycle
- Unit (dipertahankan)
- Operator (dipertahankan)
- Start (diformat ulang sebagai MM/DD/YYYY H:MM:SS)
- Finish (diformat ulang sebagai MM/DD/YYYY H:MM:SS)
- Dur (durasi dihitung ulang berdasarkan Start dan Finish jika memungkinkan)
- Source (dipertahankan)
- Destination (dipertahankan)

### Sheet Performance
- Unit (unit unik dari sheet Delay dan Cycle)
- Operating hrs (total durasi cycle per unit)
- Delay (total durasi kategori DELAY)
- Extended (total durasi kategori EXTENDED)
- Planned Down (total durasi kategori PLANNED DOWN)
- Standby (total durasi kategori STANDBY)
- Unplanned Down (total durasi kategori UNPLANNED DOWN)
- Grand Total (selalu 24 jam)
- PA (Performance Availability, persentase dengan formula: ((Grand Total - Delay) - SUM(Extended, Planned Down, Standby, Unplanned Down))/(Grand Total - Delay) * 100)

## Teknologi

- Flask (backend web)
- Pandas (pemrosesan data)
- OpenPyXL (formatting Excel)
- HTML/CSS/JavaScript (frontend)

© 2025 MINESTAR REPORTING. TECHNOLOGY DEPT.
