# Enhanced Excel Cleaner

Enhanced Excel Cleaner adalah program Python yang memungkinkan Anda membersihkan berkas Excel dengan berbagai cara, termasuk memutus tautan eksternal, menghapus lembar tersembunyi, dan menghapus nama yang salah dalam berkas Excel.

## Persyaratan

Pastikan Anda memiliki Python terinstal di sistem Anda. Kemudian, Anda dapat menginstal semua dependensi yang diperlukan dengan menjalankan perintah berikut di terminal:

```bash
pip install -r requirements.txt
```

Penggunaan
Clone atau unduh repositori ini ke komputer Anda.

Pastikan Anda berada di direktori tempat program berada.

Jalankan program dengan perintah berikut:

```bash
python cleanexcel.py
```

Ikuti petunjuk yang muncul di layar untuk memasukkan path ke berkas Excel yang ingin Anda bersihkan dan pilih tindakan yang ingin dilakukan.

Opsi Tindakan
Program ini menawarkan beberapa opsi tindakan:

1. Putuskan tautan eksternal (win32com)
2. Hapus lembar tersembunyi (openpyxl)
3. Hapus nama yang salah (win32com)
4. Lakukan semua tindakan di atas

Cukup masukkan angka yang sesuai dengan tindakan yang ingin Anda lakukan dan program akan mengeksekusinya untuk Anda.

Contoh
Setelah menjalankan program, Anda akan melihat petunjuk seperti ini di terminal:

```plaintext
Welcome to Enhanced Excel Cleaner
Enter the path to the Excel file: [Masukkan path ke berkas Excel]
Choose an action:
1. Break external links (win32com)
2. Delete hidden sheets (openpyxl)
3. Delete erroneous names (win32com)
4. Perform all actions
Enter your choice (1-4):
```
Masukkan pilihan yang diinginkan dan tunggu hingga proses selesai.


Selamat membersihkan berkas Excel Anda!
