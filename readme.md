# README: VBA Script untuk Input Field SPOP-LSPOP di Excel

## Deskripsi

Script VBA ini dirancang untuk memproses data mentah (raw data) di Microsoft Excel dengan memudahkan input data ke dalam field yang telah ditentukan.

## Fitur Utama

- Memproses data mentah dari sumber eksternal atau sheet lain
- Memvalidasi input data sebelum diproses
- Memindahkan data ke lokasi yang ditentukan (soon)
- Mengotomatisasi proses input data untuk menghemat waktu

## Prasyarat

- Microsoft Excel (2010 atau yang lebih baru)
- Macro dan VBA diaktifkan
- Izin untuk menjalankan macro di dokumen Excel

## Cara Menggunakan

1. **Buka File Excel**:

   - Buka file Excel yang berisi script VBA ini

2. **Aktifkan Developer Tab**:

   - Buka File > Options > Customize Ribbon
   - Centang opsi "Developer" dan klik OK

3. **Akses VBA Editor**:

   - Klik tab Developer
   - Pilih "Visual Basic" atau tekan `Alt + F11`

4. **Import Script**:

   - Jika script belum ada, impor module baru dan salin kode VBA ke dalamnya

5. **Jalankan Macro**:
   - Kembali ke Excel
   - Tekan `Alt + F8`, pilih macro yang sesuai, dan klik "Run"

## Parameter Input

Script ini menerima input dari:

- Range tertentu di worksheet
- File eksternal (sesuai konfigurasi script)
- Form input khusus (jika disediakan)

## Customisasi

Anda dapat menyesuaikan:

- Range sumber data dengan mengubah variabel `sourceRange`
- Range tujuan dengan mengubah variabel `targetRange`
- Aturan validasi dengan memodifikasi fungsi validasi

## Troubleshooting

- **Macro tidak berjalan**: Pastikan macro diaktifkan di Trust Center Settings
- **Data tidak terbaca**: Periksa format data sumber dan sesuaikan dengan yang diharapkan script
- **Error runtime**: Periksa Debug untuk melihat baris mana yang menyebabkan error

## Kontribusi

Jika Anda ingin berkontribusi pada pengembangan script ini, silakan:

1. Fork repository ini
2. Buat branch untuk fitur baru (`git checkout -b fitur-baru`)
3. Commit perubahan Anda (`git commit -am 'feat: ...'`)
4. Push ke branch (`git push origin new-feature`)
5. Buat Pull Request

## Lisensi

Script ini dilisensikan di bawah [MIT License](LICENSE).

---

Untuk pertanyaan lebih lanjut, silakan hubungi pengembang.
