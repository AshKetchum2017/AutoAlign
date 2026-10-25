# Roadmap AutoAlign Maximum

Tanggal perencanaan: **24 Oktober 2026**.

Dokumen ini mengatur prioritas implementasi Maximum beserta daftar kandidat dan Apply ulang. **Seluruh scope implementasi di bawah masih terbuka.** Source saat penyusunan roadmap menyediakan Minimum/Medium; `optMaximum` masih dinonaktifkan.

## 1. Tujuan dan aturan tetap

Maximum mengutamakan jumlah objek dan pemanfaatan ruang yang layak diisi. Kerapian tidak menjadi syarat. Pencarian tetap berbatas dan tidak menjanjikan nesting optimum global.

### Quantity dan pemilihan Design

| Input | Aturan Maximum |
| --- | --- |
| Angka positif | Penuhi jumlah tepat. Utamakan jumlah container paling sedikit untuk memenuhi kebutuhan tersebut. |
| `0` | Lewati Design. |
| Kosong / `-` | Boleh ditambahkan sebagai pengisi ruang. Design kecil boleh mendominasi, dan Design tertentu boleh tidak dipakai demi total objek terbanyak. |
| `/n` | Satu container per kelompok; kelompok tetap terpisah dan pembagian jumlah antar-Design dijaga dengan selisih maksimal satu. |
| `/n 80`, `/n:80`, `/n=80` | Target 80 per Design. Pertahankan prioritas target kelompok tujuan, penerusan sisa dan container tambahan sesuai kontrak Quantity. |
| Suffix `*` | Tolak objek kelompok lain pada container tersebut; sisa target sendiri boleh diteruskan. |
| Suffix `&` | Boleh menarik filler tanpa Quantity tetap dari kelompok berikutnya. Jangan menarik target angka. Tidak digabung dengan `*`. |

- Boleh memakai Design lain yang terdaftar untuk mengisi celah, sesuai konteks container dan aturan Quantity. Perpindahan lintas kelompok mengikuti aturan sisa target dan modifier; jangan mencampur kelompok di luar aturan tersebut. Pembagian rata berlaku pada anggota tanpa target tetap, sedangkan target angka dihitung per Design.
- Boleh menggeser objek dan mengganti orientasinya selama perencanaan.
- Pasangan Design/Cut Line, template bersama, gap, mark, batas area dan margin Custom tetap harus valid.
- Input mengikuti pemetaan pasangan yang didukung AA. Beberapa kelompok berbagi Cut Line berbeda masuk perbaikan prasyarat Process di bawah. Campuran standalone/berpasangan tetap di luar scope; Optimize/Periksa Susunan untuk beberapa kelompok Cut bersama dijalankan per kelompok.
- Perubahan Maximum tidak mengubah kontrak Minimum, Medium, Optimize, atau Periksa Susunan.

### Rotasi

- Khusus Maximum, seluruh preset `cmbRotation` menjadi **acuan sudut awal** untuk refinemen lokal.
- Termasuk **Tanpa rotasi**: acuan awal 0° boleh disempurnakan di sekitarnya.
- Contoh yang disepakati: `90° → 85–95° → 87° → 86–88° → 86,8° → 86,7–86,9°`.
- Radius, step, jumlah level dan anggaran pencarian belum ditetapkan. Angka contoh bukan penetapan seluruh parameter produksi.
- Sudut dinilai berdasarkan susunan valid yang dihasilkan, bukan hanya ukuran bounding box satu objek.

### Kandidat dan Apply ulang

- Simpan beberapa kandidat valid dan berbeda dengan jumlah terbaik yang sama, dalam konteks kebutuhan container/Quantity yang sama.
- Kandidat dengan jumlah lebih tinggi menggantikan kandidat berjumlah lebih rendah pada konteks tersebut. Jumlah lebih tinggi tidak boleh melanggar Quantity atau menambah container bebas tanpa batas.
- User memilih kandidat melalui ListBox, lalu menjalankan Apply.
- Setelah Apply, kandidat lain boleh diterapkan untuk **mengganti hasil AA sebelumnya**.
- Cache hanya selama sesi AA. Menutup AA membuang cache, tetapi mempertahankan hasil terakhir yang sudah diterapkan.

## 2. Prioritas dan dependensi

| Urutan | Scope | Prioritas | Bergantung pada | Hasil yang dituju |
| --- | --- | --- | --- | --- |
| S0 | Kontrak job, penilaian dan acuan regresi | P0 | — | Makna input, skor dan batas ownership jelas. |
| S1 | Snapshot sumber dan kandidat | P0 | S0 | Sumber awal stabil dan kandidat tidak saling berubah. |
| S2 | Apply dan penggantian hasil | P0 | S1 | Apply A → B aman terhadap Undo, Cancel dan perubahan luar. |
| S3 | Daftar kandidat dan jalur dasar Maximum | P0 | S2 | Alur Search → pilih → Apply dapat diuji memakai hasil dasar yang valid. |
| S4 | Pengisian celah lintas Design | P1 | S3 | Maximum dasar menambah objek dengan Quantity tetap benar. |
| S5 | Beberapa susunan awal dan refinemen X/Y | P1 | S4 | Pencarian tidak bergantung pada satu pola atau posisi awal. |
| S6 | Refinemen sudut bertingkat | P1 | S5 | Sudut desimal aman dan konsisten di semua jalur Maximum. |
| S7 | Pergeseran kelompok kecil | P2 | S6 | Ruang tersisa dapat disatukan tanpa merusak pasangan. |
| S8 | Lepas dan isi ulang bagian lokal | P2 | S7 | Susunan yang mentok dapat diperbaiki melalui pergantian lokal. |
| S9 | Pengukuran dan penerimaan keseluruhan | P0 pada penyelesaian | S4–S8 | Cakupan dan biaya diketahui sebelum dinyatakan lengkap. |

**P0** adalah fondasi atau gerbang keamanan. **P1** adalah kemampuan utama. **P2** adalah pencarian lanjutan dengan biaya lebih besar.

S9 menggabungkan hasil, tetapi pengukuran dan pengujian tetap dilakukan pada setiap scope. Tidak perlu mengimplementasikan seluruh roadmap dalam satu perubahan.

## 3. Scope implementasi

### Prasyarat P0 — Regresi Process beberapa kelompok Cut bersama

Kasus sumber: Design `/81`, `/80`, `/79` berbagi Cut `/32`, dan Design `/41` memakai Cut `/78`. Perbaikan source Process tersedia. Tutup matriks regresi berikut sebelum fondasi Maximum dilanjutkan; bukti dan cakupan setiap pengujian dicatat di `Development.log` workspace.

- [x] Uji Minimum dan Medium: reuse template per Cut untuk Quantity biasa, termasuk Design pertama Quantity 0.
- [x] Uji `/n` lintas Cut dan kombinasi kelompok `/n` dengan Quantity biasa; pembagian jumlah dan pemisahan container tetap benar.
- [x] Uji target per Design `/n 80`, penerusan sisa, filler, `*` dan `&`, termasuk container tambahan serta pasangan Cut yang ikut berpindah.
- [x] Periksa jumlah Cut fisik, rotasi per kelompok, gap, mark, bounds, laporan, Cancel dan satu Undo.
- [x] Pastikan input lama tetap bekerja dan Optimize/Periksa Susunan menolak struktur di luar cakupannya sebelum mutasi.

**Gerbang lanjut:** sumber multi-Design/per-Cut menjadi acuan regresi yang valid untuk S0. Status dan bukti berada di `Development.log` workspace; perbaikan ini tidak mengaktifkan Maximum.

### S0 — Kontrak job dan acuan regresi

- [ ] Tetapkan identitas job: dokumen/halaman, sumber dan pasangan, Quantity, preset rotasi, model, area, gap, margin serta zona mark.
- [ ] Tetapkan penilaian untuk Quantity tetap, pengisian bebas dan `/n`. Pisahkan syarat validitas dari pemeringkatan.
- [ ] Gunakan kandidat **job lengkap** sebagai unit pilihan awal, termasuk container kelompok dan ekspansi template. Pemilihan per-container dapat dikaji kemudian.
- [ ] Tentukan pengaturan yang membatalkan cache serta cara mengenali perubahan sumber/output di luar AA.
- [ ] Pilih dokumen uji dan catat acuan jumlah, bounds, rotasi, keamanan serta waktu sebelum perubahan.

**Kriteria selesai:** aturan penilaian dan konteks yang boleh dipakai ulang dapat diuji; tidak ada ambiguitas yang memengaruhi jumlah wajib, kelompok atau penghapusan hasil.

### S1 — Snapshot sumber dan kandidat yang independen

`AAResultBuilder.TakeShape` saat ini memakai sumber asli sebagai instance pertama. Sumber tersebut dapat masuk ke group hasil. Karena itu, Apply ulang memerlukan sumber awal yang tetap tersedia.

- [ ] Tentukan representasi baseline/template sumber per job berdasarkan kemampuan host. Hindari salinan artwork untuk setiap kandidat.
- [ ] Simpan orientasi dan data sumber awal agar pergantian kandidat tidak menumpuk delta rotasi dari hasil sebelumnya.
- [ ] Snapshot kandidat memuat identitas sumber/pasangan, area, posisi, sudut, gap, suppression/template dan penanda allowance yang diperlukan untuk rekonstruksi.
- [ ] Pisahkan snapshot dari placement mutable planner. `CopyForPlan` dapat menjadi referensi pola penyalinan, tetapi tetap menahan geometry/contour yang dibagikan.
- [ ] Ikat resource sementara ke sesi dan job. Bersihkan pada pembatalan, pergantian job dan penutupan AA tanpa menghapus hasil yang telah diterapkan.

**Kriteria selesai:** mengubah atau memusatkan kandidat B tidak mengubah A; baseline tetap tersedia setelah Apply; tidak ada resource/objek sementara tertinggal. Ukur memori baseline terpisah dari metadata kandidat, terutama pada bitmap.

### S2 — Apply ulang sebagai transaksi penggantian

- [ ] Pisahkan persiapan/pencarian rencana dari penerapan ke dokumen.
- [ ] Resolve sumber dan identitas output milik job sebelum mutasi. Jangan memakai selection saat ini sebagai daftar hasil yang boleh dihapus.
- [ ] Validasi kandidat serta konteks sumber/pengaturan sebelum Apply, dan periksa hasil aktual setelah dibuat.
- [ ] Ganti hanya output job tersebut. Pertahankan hasil sebelumnya bila pembuatan atau pemeriksaan kandidat baru gagal/Cancel.
- [ ] Satu Apply menjadi satu CommandGroup. Jangan menggunakan Undo global sebagai mekanisme mengganti kandidat.
- [ ] Publikasikan laporan produksi setelah commit dan pemulihan state berhasil; kegagalan mempertahankan laporan sukses sebelumnya.
- [ ] Tangani Undo/Redo, edit luar dan ID hasil yang hilang: validasi ulang ownership atau tolak pergantian, tanpa menghapus pekerjaan pengguna.

**Kriteria selesai:** Apply A → B, Undo → A dan Redo → B benar; kegagalan/Cancel mempertahankan keadaan yang ditentukan. Unit, ruler, halaman, layer, printable state dan selection dipulihkan sesuai kontrak.

### S3 — Daftar kandidat dan jalur dasar Maximum

- [ ] Tempatkan pool kandidat pada pemilik sesi form/presenter, bukan variabel global yang bertahan setelah AA ditutup.
- [ ] Simpan kandidat valid yang berbeda dan berjumlah terbaik sama. Deduplikasi memperhitungkan identitas, posisi, sudut, area dan struktur pasangan/template; ukuran bbox saja tidak cukup.
- [ ] Batasi pool melalui `Private Const`. **Lima kandidat** adalah usulan awal eksperimental, bukan angka final.
- [ ] Jangan memperpanjang pencarian hanya untuk memenuhi kapasitas pool. Tampilkan kandidat yang memang tersedia.
- [ ] Rencana kontrol: `lbxCandidates` dan `cmdApplyCandidate`, terpisah dari `lbxObjects`; informasi jumlah, container, bounds dan komposisi/orientasi ditampilkan untuk membantu pemilihan.
- [ ] Memilih baris tidak mengubah dokumen. Preview ringan bersifat opsional untuk tahap berikutnya.
- [ ] Hubungkan alur Maximum ke fondasi tersebut memakai hasil awal yang valid dari Medium. Pada tahap ini, refinemen sudut belum tersedia.

**Kriteria selesai:** user dapat memilih dan mengganti kandidat tanpa pencarian ulang; perubahan konteks membatalkan cache yang tidak lagi sesuai. Aktivasi untuk uji tahap dasar tidak berarti seluruh Maximum selesai. Minimum/Medium mempertahankan alur yang ada.

### S4 — Maximum dasar: pengisian celah lintas Design

- [ ] Mulai dari kandidat Medium yang valid dan pertahankan sebagai pembanding.
- [ ] Isi ruang layak dengan Design yang tersedia dalam konteks container yang diizinkan.
- [ ] Lindungi jumlah wajib dan pembagian `/n` sebelum menambahkan filler bebas; jangan membuat jumlah wajib berlebih.
- [ ] Nilai urutan Design wajib yang sulit ditempatkan, tanpa mengunci seluruh pencarian pada satu urutan.
- [ ] Pakai pemeriksaan kontur, gap, overlap, area dan mark yang berlaku. Jangan mengorbankan validitas demi menambah objek.
- [ ] Masukkan hasil akhir valid ke pool. Batasi pass/probe dan tampilkan progres serta checkpoint Cancel.

**Kriteria selesai:** ruang kosong dapat diisi Design lain tanpa kebocoran kelompok, perubahan jumlah wajib atau kerusakan pasangan/template. Hasil awal tetap tersedia bila upaya tambahan tidak memperbaiki skor.

### S5 — Variasi susunan awal dan refinemen posisi

- [ ] Coba sejumlah terbatas arah mulai, urutan Design dan pola susunan yang berbeda.
- [ ] Simpan beberapa alternatif internal yang berbeda, termasuk seri jumlah yang memiliki peluang perbaikan. Pool internal pencarian dapat berbeda dari pool yang ditampilkan ke user.
- [ ] Sempurnakan posisi X/Y secara lokal untuk membuka celah, dengan langkah dan jumlah evaluasi terbatas.
- [ ] Hindari membangun seluruh kombinasi posisi × sudut sekaligus. Siapkan proses bergantian posisi dan sudut untuk S6.

**Kriteria selesai:** jumlah cabang dan biaya dapat diprediksi dari budget; hasil dapat direproduksi. Jika memakai pengacakan, seed dicatat di debug. Kandidat awal terbaik tidak hilang.

### S6 — Refinemen sudut bertingkat

- [ ] Terapkan kebijakan khusus Maximum: preset adalah acuan awal, termasuk penyempurnaan di sekitar 0° untuk Tanpa rotasi.
- [ ] Selaraskan kebijakan pada `AAModel.AllowsRotation`/`RotationVariant`, pasangan dan irisan Cut bersama, template/cache, serta validasi akhir. Jangan melonggarkan aturan global Minimum/Medium.
- [ ] Mulai dari sudut awal yang menjanjikan lalu persempit rentang, mengikuti contoh 90 → 87 → 86,8 derajat.
- [ ] Putuskan budget awal dalam `Private Const`: radius/step, level, jumlah sudut awal, simpangan total, evaluasi per Design/per run dan varian geometri tersimpan.
- [ ] Rekomendasi: jepit refinemen dalam jendela awal dan normalisasi batas 0°/360°; jangan mengevaluasi sudut identik berulang.
- [ ] Lakukan refinemen posisi dan sudut secara bergantian. Pertahankan hubungan Design/Cut serta pola pasangan berlawanan bila digunakan.
- [ ] Nilai susunan valid yang dihasilkan. Sudut dengan bbox objek lebih kecil belum tentu memberi jumlah lebih banyak.
- [ ] Perbarui hint agar makna preset khusus Maximum jelas.

**Kriteria selesai:** sudut desimal dapat dicari dan diterapkan tanpa akumulasi rotasi, pelanggaran preset Maximum atau perubahan perilaku mode lain. Batas budget dan hasil valid teruji.

### S7 — Pergeseran kelompok kecil

- [ ] Pilih kelompok lokal dengan ukuran terbatas dan coba gerak bersama untuk menyatukan ruang.
- [ ] Pertahankan hubungan internal yang diperlukan; periksa kembali seluruh hubungan dengan tetangga, mark dan area.
- [ ] Coba isi ruang yang terbentuk dan bandingkan hasil lengkap dengan kandidat sebelumnya.

**Kriteria selesai:** pergeseran menghasilkan perbaikan skor yang valid atau dikembalikan; tidak ada drift seluruh susunan tanpa batas.

### S8 — Lepas dan isi ulang lokal

- [ ] Lepas sementara beberapa placement dari rencana, lalu coba isi ulang kantong tersebut dengan susunan berbeda.
- [ ] Untuk pengisian bebas, pertukaran satu objek dengan dua objek lain boleh dicoba. Semua jumlah wajib dan pembagian kelompok tetap harus terpenuhi pada hasil akhir.
- [ ] Batasi objek yang dilepas, luas lingkungan, cabang dan percobaan per kantong.
- [ ] Jangan memutasi Shape dokumen selama trial. Hasil kalah/invalid tidak menggantikan kandidat terbaik.

**Kriteria selesai:** pencarian lokal tidak menghilangkan kebutuhan wajib dan tidak menurunkan hasil terbaik; kegagalan atau habis budget mengembalikan kandidat valid yang tersedia.

### S9 — Pengukuran dan penerimaan keseluruhan

- [ ] Jalankan matriks regresi di bawah pada cakupan yang telah diimplementasikan.
- [ ] Ukur Search, validasi, Apply/penggantian, metadata pool dan baseline native secara terpisah.
- [ ] Catat jumlah kandidat, level/evaluasi sudut, perbaikan yang diterima dan alasan berhenti.
- [ ] Lakukan perbandingan berulang dengan sumber/pengaturan identik sebelum mengklaim percepatan umum.
- [ ] Perbarui status fitur dan dokumentasi berdasarkan cakupan yang benar-benar selesai. Sisakan scope lanjutan yang belum dikerjakan secara eksplisit.

**Kriteria selesai:** aturan yang disepakati memiliki bukti sesuai cakupannya, keterbatasan jelas, dan tidak ada regresi mode lain yang belum ditangani.

## 4. Matriks pengujian

| Kelompok | Kasus representatif | Yang diperiksa |
| --- | --- | --- |
| Bentuk | Rectangle, rounded rectangle, ellipse, concave, PinPoint, kurva kompleks | Jumlah, ruang layak, rotasi, gap dan collision. |
| Sumber | Standalone, bitmap + Cut, satu Cut bersama, Cut per Design | Identitas, pasangan, template dan suppression. |
| Quantity | Angka, 0, kosong, beberapa `/n`, kombinasi yang valid | Jumlah tepat, filler dan isolasi/balance kelompok. |
| Area | KissA, DieA, Custom dengan margin/zona | Mark, bounds, aturan allowance yang berlaku dan grid area. |
| Rotasi | Semua preset, sudut sumber awal, sekitar 0°/360°, sudut desimal | Kebijakan Maximum, referensi sudut awal dan regresi Minimum/Medium. |
| Apply ulang | A → B → A, Undo/Redo, ganti kandidat setelah perubahan luar | Hanya output job yang berubah; sumber, ownership dan laporan tetap benar. |
| Kegagalan | Cancel Search/Apply, ID hilang, gagal build/validasi | Tidak ada hasil parsial, sumber hilang atau output lama yang terhapus tanpa pemulihan. |
| Sesi | Tutup AA sebelum/sesudah Apply, job baru, pengaturan berubah | Cache/resource dibersihkan; hasil terapan tetap ada; kandidat usang ditolak. |
| State dokumen | Unit/ruler, halaman/layer, printable state, selection | Pemulihan sesuai kontrak pada sukses dan kegagalan. |

Pemeriksaan statis dan harness adaptasi membantu menguji kontrak, tetapi **tidak menggantikan import Class Module, compile VBE, runtime CorelDRAW dan pemeriksaan manual**. Catat bukti serta cakupannya di `Development.log` workspace.

## 5. Batas scope dan cara melanjutkan

- Kerjakan satu scope dan tutup risiko konkretnya sebelum memperluas pencarian.
- Tutup prasyarat regresi Process beberapa kelompok Cut bersama, lalu lanjutkan **S0–S1: kontrak job, snapshot dan baseline sumber**. Apply ulang bergantung pada fondasi ini.
- Untuk perubahan kode, periksa ulang active source, caller dan log terkait. Nama kontrol di roadmap adalah rencana; resource form belum diubah.
- Optimasi MediumPlanner lanjutan yang ditunda tidak otomatis dilanjutkan oleh roadmap Maximum.
- Migrasi ke AutomatonStudio, penyimpanan kandidat lintas sesi, preview penuh artwork dan pencarian tanpa batas bukan bagian implementasi awal.
- Pelihara progres scope di dokumen ini. Simpan riwayat keputusan/pengujian di empat log root sesuai perannya; jangan menandai fitur selesai hanya karena dokumen rencananya selesai.

## 6. Acuan source

| Komponen aktif | Peran yang perlu diperiksa saat implementasi |
| --- | --- |
| `src/forms/AutoAlignMenu.vba` | Mode, kontrol, progres, Cancel dan lifecycle form. |
| `src/classes/AAPresenter.cls` | Ownership sesi/job, daftar kandidat dan laporan hasil. |
| `src/classes/AAObjectStore.cls` | Identitas sumber, Quantity, rotasi dan hubungan pasangan. |
| `src/classes/AAProcessEngine.cls` | Persiapan, routing planner, template/kelompok, validasi dan transaksi hasil. |
| `src/classes/AAQuantityFlow.cls`, `AAQuantityPlanner.cls` | Kontrak target per Design, sisa lintas container, modifier dan alokasi satu area. |
| `src/classes/AAResultBuilder.cls` | Penggunaan sumber, duplikasi, penempatan dan grouping. |
| `src/classes/AALayoutPlan.cls`, `AAPlacement.cls` | Data rencana, snapshot dan metadata placement. |
| `src/classes/AAModel.cls`, `AARotationVariant.cls` | Kebijakan/varian orientasi dan geometri sumber. |
| `src/classes/AAPlanner.cls`, `AAMediumPlanner.cls` | Rencana awal, penilaian, slot, perapatan serta validator yang dapat digunakan. |
| `src/classes/AASettings.cls`, `src/modules/AAAlgorithm.bas` | Kontrak mode/rotasi, gap dan area. |

Peta ini bukan kewajiban mengubah semua komponen sekaligus. Tentukan perubahan terkecil yang menuntaskan tiap scope.
