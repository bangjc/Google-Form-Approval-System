# Google-Form-Approval-System
Sistem otomatisasi alur kerja (workflow) untuk persetujuan (approval) internal. Sistem ini terintegrasi dengan Google Form sebagai form pengajuan, Google Sheets sebagai database utama dan menggunakan Google Apps Script untuk mengirimkan email notifikasi berjenjang (ACC1>ACC2>ACC3) serta pembuatan dokumen PDF otomatis setelah persetujuan selesai.

# Fitur Utama
1. Workflow Berjenjang: Notifikasi otomatis kepada approver sesuai level (ACC1, ACC2, ACC3).
2. Approval Link: Tombol interaktif (Approve/Decline) langsung di dalam email.
3. PDF Generation: Pembuatan dokumen memo internal dalam format PDF menggunakan Google Docs sebagai template.
4. Auto-Logging: Pencatatan waktu (timestamp) otomatis setiap kali ada perubahan status.

# Teknologi yang Digunakan
1. Google Apps Script (GAS)
2. Google Sheets API
3. Google Drive API (untuk manajemen PDF)
4. MailApp Service

# Cara Instalasi
1. Buat Google Sheet baru dengan struktur kolom yang sesuai (sesuaikan dengan kode Anda).
2. Siapkan Template dokumen di Google Docs dan catat Template ID dan Folder ID.
3. Salin skrip dari file Code.gs ke Apps Script Editor di Google Sheet Anda.
4. Update konfigurasi di bagian atas file Code.gs (ID Spreadsheet dan URL Web App).
5. Deploy sebagai Web App (akses: "Anyone", jalankan sebagai: "Me").
6. Gunakan URL yang dihasilkan untuk memperbarui variabel WEB_APP_URL di skrip.

# Struktur File
- Code.gs: Logika utama backend untuk alur approval dan generate PDF.
- Template_CreditLimit.html: Struktur email yang dikirimkan kepada approver.
- File Google Docs untuk format file dan sebagai dasar untuk generate file ke Pdf.
