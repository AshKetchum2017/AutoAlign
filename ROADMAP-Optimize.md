# Roadmap Optimize dan Refill AutoAlign

Tanggal baseline perencanaan: **3 November 2026**.

Dokumen ini mengatur pengembangan tombol **Optimize** serta rencana **Refill pada susunan yang sudah ada**. Kontrak awal dan prioritas sudah disepakati; implementasi perbaikan dan pengukuran native masih terbuka. Baseline perencanaan bukan baseline waktu atau bukti bahwa algoritme sudah diperbaiki.

Prioritas, pendekatan, dan kasus uji dapat ditinjau ulang berdasarkan bukti. Perubahan kontrak jumlah, rotasi, sumber pengisi, atau area perlu keputusan tersendiri; perubahan urutan pekerjaan tidak otomatis mengubah kontrak tersebut.

## 1. Batas pekerjaan dan dokumen terkait

| Pekerjaan | Tujuan | Acuan |
| --- | --- | --- |
| Optimize Minimum/Medium | Memperbaiki susunan yang ada dengan jumlah dan rotasi objek tetap | Dokumen ini |
| Refill container | Menambah objek dari Design yang sudah ada dalam container target, dengan pergeseran lokal yang diizinkan | Dokumen ini |
| Optimasi performa Process | Mempercepat planner dan helper Process dengan kontrak hasil terjaga | [ROADMAP-Optimization.md](ROADMAP-Optimization.md) |
| Planner/algoritme Maximum | Membangun kemampuan Process Maximum dan fondasi terkait | [ROADMAP-Maximum.md](ROADMAP-Maximum.md) |

Refill Process yang sudah tersedia menjadi referensi implementasi, bukan bukti tersedianya operasi Refill untuk container terdaftar. Tahap ini tidak mengaktifkan Maximum, mengubah planner Process, atau memigrasikan fitur ke AutomatonStudio. Rancangan kontrol/form Refill belum ditetapkan.

## 2. Baseline source dan masalah awal

Acuan inspeksi awal: commit **`26dfb73`**. Sebelum implementasi, cocokkan kembali source workspace dan source yang dimuat di VBE; commit ini tidak membekukan revisi pekerjaan berikutnya.

| Komponen | Perilaku source saat baseline | Fokus pemeriksaan |
| --- | --- | --- |
| `AAOptimizeEngine.Run` | Membaca container dari Objects, memilih Minimum/Medium, memvalidasi rencana dan hasil aktual, serta mengelola Apply/pemulihan | Integritas operasi dan biaya tiap tahap |
| `AAOptimizeMinimum.Align` | Snapshot baris/kolom dari posisi awal; meratakan baris lalu kolom; toleransi terhadap anchor sebesar 25% tinggi/lebar | Pengenalan dan pelestarian fase ZigZag |
| `AAOptimizeEngine.Compact` | Menggeser unit baris/kolom/individual pada empat arah, menilai luas bounds dan sebaran, lalu mencoba centering | Biaya pencarian serta mutu hasil lokal |
| `AAOptimizeLayout` | Membaca anggota, mencocokkan template, menyimpan placement dan validasi gap/area/mark | Konsistensi snapshot, pasangan, area dan penerapan |

Parameter awal yang perlu dicatat ketika berubah: Minimum maksimal **200000 probe**; Medium maksimal **3 pass / 20000 probe**, langkah awal **1 mm**, dan ambang berhenti **0,001 mm**. Toleransi pencocokan template **1 mm** berbeda dari toleransi gap dan deteksi pola; jangan menyamakan penggunaannya.

Masalah yang menjadi sasaran awal:

- **Minimum:** pola ZigZag belum terdeteksi dengan benar dalam Optimize. Mekanisme baris/kolom saat ini belum memodelkan fase ZigZag secara khusus. Interaksi tahap baris dan kolom perlu diuji sebelum menetapkan akar masalah kasus tertentu.
- **Medium:** pengujian masih terbatas; biaya satu Optimize dirasakan lebih lama daripada menyusun ulang melalui Process Medium. Belum tersedia baseline waktu berulang atau profil yang menetapkan tahap dominan. Process menghasilkan susunan baru, sedangkan Optimize mempertahankan instance dan rotasi; kecepatan tidak dibandingkan tanpa mencatat perbedaan hasil/kontrak tersebut.
- **Maximum:** Optimize Maximum menunggu kesiapan planner dan algoritme Maximum. Kontrak pergeseran/rotasi khususnya belum ditetapkan; aturan refinemen sudut Process Maximum tidak otomatis berlaku pada Optimize.

Angka jumlah, bounds, koordinat, dan waktu kasus gagal belum ditetapkan di dokumen ini. Isi melalui kasus uji yang dapat direproduksi, tanpa mengarang hasil dari pemeriksaan source atau screenshot.

## 3. Kontrak Optimize yang dipertahankan

- Target tetap berasal dari container yang didaftarkan ke `lbxObjects`, termasuk susunan manual yang mengikuti struktur AA. Selection saat ini tidak mengganti target terdaftar.
- Struktur acuan `container → subgroup sumber → instance`; group pembentuk satu instance tetap utuh. Jumlah dan rotasi instance lama dipertahankan.
- Design dan Cut Line bergerak bersama dengan offset pasangan tetap. Template Cut bersama memerlukan susunan relatif yang cocok, dengan toleransi pencocokan eksperimental 1 mm di `Private Const`.
- Dukungan input saat baseline: standalone, satu Cut bersama, pasangan satu-per-satu, dan Cut Line saja. Campuran standalone/pasangan serta beberapa kelompok Cut bersama dalam satu Optimize belum otomatis didukung hanya karena Process mendukungnya.
- Gap dan Area mengikuti pengaturan; model menentukan Default serta mark. Quantity tidak menambah atau mengurangi objek pada Optimize.
- Area umumnya berpusat pada bounds container saat dibaca. Jika pembatas Custom aktif, source memakai grid Process tetap untuk margin/zona. Perbedaan terhadap acuan Periksa Susunan harus tetap terlihat, bukan diubah diam-diam.
- Gap, collision, mark, bounds dan margin/zona Custom wajib valid. Minimum mempertahankan anchor tanpa centering ulang; Medium mencoba centering jika hasilnya valid. Perubahan kebijakan anchor/centering memerlukan scope tersendiri.
- Rencana dibuat sebelum penerapan dan hasil aktual diperiksa kembali. Cancel, pemulihan kegagalan, satu CommandGroup, Undo/Redo, unit/ruler, selection, halaman/layer serta state dokumen tetap menjadi bagian penerimaan.

Tidak setiap pasangan harus mempunyai jarak persis nilai Gap: kontur, tetangga lain, mark dan area dapat menahan pergeseran. Klaim gap berlebih perlu posisi pembanding yang benar-benar valid.

## 4. Kontrak awal Refill

Keputusan yang sudah ditetapkan:

- Sumber tambahan **hanya Design yang sudah ada di container target**. Tidak mengambil Design lain dari ListBox atau container lain sebagai filler.
- Objek lama **boleh digeser secara lokal** untuk membuka celah, dengan rotasi lama tetap.
- Objek tambahan **boleh memakai orientasi menurut `cmbRotation`**; tidak dibatasi pada orientasi yang sudah muncul. Cara memperoleh preset per sumber dari struktur container masih perlu dirancang.
- Pairing, template Cut, gap, mark, area dan struktur group harus tetap valid. Penambahan Design berpasangan memerlukan Cut yang sesuai dan kebijakan suppression/template yang benar.

Usulan penerimaan: gerak lokal diterima bersama penambahan objek valid; percobaan yang gagal mempertahankan susunan sebelumnya. Finalisasi aturan ini termasuk tahap B4, bukan alasan melakukan pergeseran tanpa batas.

Keputusan yang masih terbuka sebelum implementasi Refill:

- Jumlah tambahan, batas total, dan pembagian ketika beberapa Design bersaing; termasuk makna Quantity yang dimasukkan pada daftar container dan hubungan dengan `/n`, `*`, `&` atau target produksi asli. Riwayat Quantity Process tidak boleh diasumsikan tersimpan pada hasil.
- Cara mengenali sumber yang sama dan pasangan tiap instance di dalam container; jangan menganggap semua instance group sebagai Design baru atau memilih pasangan hanya dari bentuk visual.
- Radius/arah/jumlah objek yang boleh digeser, budget probe/pass/waktu, aturan berhenti dan ranking kandidat. Simpan batas eksperimental di `Private Const`; angka belum ditetapkan.
- Cara menangani template bersama di beberapa container Design, termasuk identitas sumber, tambahan yang konsisten, dan satu Cut fisik yang digunakan bersama.
- Tahap operasi: Refill mandiri atau pilihan setelah Optimize, acuan area, perilaku untuk susunan awal tidak valid, kontrol UI dan pemulihan tambahan objek saat gagal/Cancel.

Refill tidak membuka pencarian sudut desimal Maximum sebelum kemampuan dan kontraknya siap.

## 5. Prioritas dan tahap pekerjaan

| ID | Scope | Prioritas/dependensi | Status awal | Kriteria selesai |
| --- | --- | --- | --- | --- |
| B0 | Kontrak dan kasus baseline | Pertama | Kontrak awal tersedia; fixture native dan pengukuran belum dikumpulkan | Kasus dapat diulang, kegagalan/hasil harapan dan revisi tercatat |
| B1 | Audit fondasi snapshot, pairing dan validasi | Sebelum perubahan algoritme terkait | Source tersedia; audit/regresi terarah masih terbuka | Anggota, pose, pasangan, area dan pemulihan terbukti pada jalur terdampak |
| B2 | Optimize Minimum: ZigZag | Sasaran perbaikan pertama, setelah B0/B1 | Belum diimplementasikan | Pola ZigZag dan fase dipertahankan, gap diperbaiki tanpa regresi Straight |
| B3 | Optimize Medium: profil lalu perbaikan biaya | Pengukuran dimulai pada B0; perubahan mengikuti profil | Belum ada baseline waktu baru | Biaya dominan diketahui; perbaikan terukur dengan hasil valid dan kontrak tetap |
| B4 | Refill dengan gerak lokal terbatas | Setelah fondasi/perilaku Optimize cukup stabil dan keputusan Refill selesai | Kontrak sebagian ditetapkan | Tambahan valid dari sumber yang diizinkan, gerak terbatas, gagal/Cancel dapat dipulihkan |
| B5 | Optimize Maximum | Menunggu kesiapan planner/algoritme Maximum dan kontrak khusus | Ditunda | Dependensi siap dan kemampuan baru diuji tanpa mengubah kontrak mode lain |

### B2 — Minimum dan fase ZigZag

- Rekam fase/offset antarbaris dari posisi awal dan gunakan Cut Line sebagai acuan bila berpasangan.
- Bedakan Straight, ZigZag, jitter dan pola campuran; jangan meratakan offset ZigZag hanya karena objek masuk ke kelompok kolom tertentu.
- Uji baris tidak lengkap, slot kosong dan kedua fase awal. Jangan mengunci pengenal ke satu diameter lingkaran atau offset setengah lebar pada semua bentuk.
- Pertahankan urutan tetangga dan jumlah, lalu periksa interaksi tahap horizontal/vertikal. Kegagalan deteksi tidak boleh diam-diam merusak pola yang sudah valid.
- Batas toleransi/heuristik serta tindakan ketika pola ambigu ditetapkan dari fixture. Tidak langsung menaikkan toleransi 25% untuk menutupi kegagalan.

### B3 — Medium dan biaya pencarian

- Pisahkan waktu persiapan/geometri, pembentukan unit gerak, probe/validasi, centering, Apply/validasi aktual, serta cleanup. Tambahkan instrumentasi hanya bila counter yang ada belum cukup.
- Rekam jumlah unit, gerakan diterima, probe, pass, alasan berhenti dan bounds sebelum/sesudah. Timer inclusive tidak dijumlahkan sebagai tahap independen.
- Pilih perubahan dari biaya yang terukur. Cache, reuse geometri atau pengurangan pemeriksaan memerlukan bukti konteks identik; tidak mengurangi validasi, budget atau kualitas secara sembarang.
- Bandingkan revisi Optimize pada input yang sama. Process Medium menjadi pembanding pengalaman penggunaan, bukan pengganti bukti A/B Optimize.

## 6. Matriks baseline dan pengukuran

| ID kasus | Susunan/kondisi | Bukti yang dibutuhkan |
| --- | --- | --- |
| OM-Straight | Baris/kolom sederhana, gap lebar/kecil, H/V berbeda termasuk 0 | Kontrol Minimum; anchor, arah koreksi, gap dan idempotensi |
| OM-ZigZag | Kasus gagal Minimum, kedua fase, baris parsial dan slot kosong | Koordinat/rotasi awal, fase, hasil aktual serta hasil manual valid |
| OD-Medium | Kasus Optimize lambat yang mewakili penggunaan | Profil tahap, jumlah/rotasi tetap, bounds, waktu berulang |
| OP-Pair | Standalone, Cut Line saja, pasangan masing-masing dan template bersama | Mapping, offset, suppression, jumlah Cut fisik serta gerak konsisten |
| OA-Area | KissA/DieA, Custom dengan/tanpa pembatas, dekat mark/batas | Acuan area, Default/override, gap/mark/bounds dan perilaku penolakan |
| OR-Refill | Celah langsung dan celah yang memerlukan gerak lokal | Sumber tambahan, pose lama, sudut baru legal, batas gerak dan tambahan valid |
| OC-State | Cancel saat pencarian/penerapan, gagal validasi, Undo/Redo | Tidak ada hasil parsial yang tidak dilaporkan; state dokumen pulih |

ID di atas adalah rencana fixture, bukan nama file uji yang sudah tersedia. Mulai dari OM-Straight, OM-ZigZag dan OD-Medium; perluas sesuai jalur yang disentuh. Implementasi tidak menunggu semua variasi selesai, tetapi penerimaan tidak melampaui kasus yang diuji.

Untuk setiap baseline:

1. Catat dokumen/halaman, ID container dan struktur sumber/Cut, revisi source/VBE, versi host, model/mode, Area/Gap, pembatas Custom, pengaturan debug dan preset rotasi bila relevan.
2. Simpan kondisi awal yang dapat dipulihkan. Run berulang memakai kondisi awal yang sama, bukan hasil Optimize sebelumnya.
3. Rekam jumlah per sumber/container, pose/rotasi, bounds, posisi pembanding manual, gap/mark, waktu per tahap dan total. Baseline yang belum diukur tetap ditandai belum tersedia.
4. Untuk waktu, mulai minimal dua pengamatan; tambah pengulangan dan gunakan A/B bergantian bila variasi mengaburkan manfaat. Pisahkan debug detail dan penggunaan normal. Belum ada target persentase wajib.
5. Bedakan inspeksi source, harness adaptasi/mock, compile VBE, runtime CorelDRAW dan pemeriksaan manual. Simpan bukti serta penerimaan pengguna di `Development.log` workspace.

## 7. Penerimaan dan pemeliharaan

Setiap tahap mencatat status direncanakan, diimplementasikan menunggu uji, diterima pada kasus tertentu, atau ditunda beserta alasannya. Penambahan roadmap ini tidak menandai B1–B5 selesai dan tidak mengesahkan Optimize untuk seluruh bentuk.

README menjelaskan fitur yang tersedia. Roadmap ini menyimpan arah dan keputusan terbuka; [Idea.log](../Idea.log) menjadi indeks tindak lanjut, sedangkan [Development.log](../Development.log) menyimpan investigasi dan bukti. Jangan menduplikasi seluruh riwayat ke tiap dokumen.

| Tanggal | Keputusan arah |
| --- | --- |
| 2026-11-03 | Baseline perencanaan dibuat: ZigZag Minimum, profiling Optimize Medium, Refill dari Design dalam container dengan gerak lokal dan sudut tambahan menurut cmbRotation; Optimize Maximum menunggu fondasi Maximum. |
