KEHADIRAN GURU - PWA WRAPPER

Tujuan
- Menjadikan Web App Google Apps Script dapat dipasang dari browser sebagai PWA.
- Tampilan aplikasi tetap berasal dari Web App Apps Script yang sudah ada.
- PWA shell memakai manifest + service worker + HTTPS hosting.

Penting
- Jangan menaruh PIN/password di repository. File ini hanya memuat URL Web App /exec.
- Aplikasi absensi tetap memerlukan internet karena backend berada di Google Apps Script.
- ikon bawaan adalah ikon sementara; ganti icon-192.png dan icon-512.png dengan logo sekolah sebelum rilis.

File
- index.html           : shell PWA + iframe ke Web App
- manifest.webmanifest : identitas, ikon, standalone
- sw.js                : service worker shell
- config.js            : tempat URL Web App /exec
- icons/               : ikon PWA
- Code.gs, Index.html  : versi Apps Script yang namanya sudah diseragamkan

Pemasangan
1. Edit config.js dan isi APP_CONFIG.appUrl dengan URL Web App /exec.
2. Upload seluruh folder ini ke hosting HTTPS, misalnya GitHub Pages.
3. Buka URL PWA melalui Chrome Android.
4. Tunggu service worker aktif.
5. Gunakan tombol 'Pasang Aplikasi' bila muncul, atau menu Chrome > Tambahkan ke layar utama.

GitHub Pages
- Buat repository baru, misalnya kehadiran-guru-pwa.
- Upload index.html, config.js, manifest.webmanifest, sw.js, dan folder icons.
- Settings > Pages > Deploy from branch > main > /(root) > Save.
- Buka URL https://NAMAUSER.github.io/kehadiran-guru-pwa/

Catatan teknis
Web App Apps Script tetap menjadi aplikasi inti di dalam iframe. Iframe diberi permission policy 'geolocation' agar fitur GPS dapat dipakai oleh halaman yang ditampilkan.
