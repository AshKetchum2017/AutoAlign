# Roadmap Optimasi AutoAlign

Dibuat: **3 November 2026**. Diperbarui: **5 November 2026**.

**Dokumen ini bersifat dinamis.** Prioritas, tahapan, kasus uji, dan pendekatan boleh diubah, dipecah, digabung, ditunda, atau dibatalkan mengikuti hasil pengukuran, perubahan source, serta kebutuhan pengguna. Urutan di bawah adalah arah awal, bukan rencana yang fixed atau janji implementasi seluruh tahap.

## Tujuan dan scope

Mempercepat Process AutoAlign dengan mempertahankan hasil dan aturan pekerjaan. Fokus awal adalah **MediumPlanner dan helper yang dipakainya**, termasuk baseline Minimum, pencarian Quantity, mark, filler, dan refill.

Optimasi performa dalam dokumen ini berbeda dari pengembangan tombol **Optimize**. Pengembangan mode **Maximum** mengikuti [ROADMAP-Maximum.md](ROADMAP-Maximum.md). Perubahan pada helper bersama perlu mempertimbangkan workflow lain yang benar-benar memanggilnya; roadmap ini tidak otomatis memperluas fitur atau memigrasikan engine ke AutomatonStudio.

Sasaran historis pengurangan waktu **60–70%** berasal dari kasus awal Case_07 yang membutuhkan sekitar 33 menit. Itu tetap aspirasi, bukan persentase yang sudah tercapai atau kewajiban setiap kasus baru. Setiap klaim penghematan harus menyebut kasus, revisi, baseline, dan kondisi pengukurannya sendiri.

## Kontrak hasil yang dipertahankan

- Jumlah dan alokasi per Design mengikuti Quantity, termasuk target tetap, filler, carry, modifier `*`/`&`, serta batas selisih kelompok yang berlaku. Percepatan tidak diperoleh dengan mengurangi hasil atau menonaktifkan keluarga kandidat secara sembarang.
- Pairing, Cut bersama, suppression template, izin rotasi, gap kontur, collision, mark, batas area, dan margin/zona Custom tetap valid.
- Kerapian, ranking kandidat, dan urutan hasil tidak boleh memburuk diam-diam. Optimasi mekanis seperti transfer data/sort harus mempertahankan urutan lengkap, termasuk tie. Perubahan heuristik memerlukan scope dan pembanding kualitas tersendiri.
- Refill mempertahankan aturan ownership dan batasnya: sumber utama yang dibekukan tetap, pengisian silang hanya pada konteks yang mengizinkannya, dan tambahan tetap melalui validator.
- Cancel, Undo/Redo, unit dokumen, selection, layer, dan pemulihan state tetap mengikuti kontrak operasi.
- Batas waktu atau pengurangan budget pencarian yang menukar kualitas dengan kecepatan tidak dianggap sekadar optimasi internal; perlu keputusan scope tersendiri.

Fleksibilitas roadmap tidak dengan sendirinya mengubah kontrak tersebut. Jika kebutuhan produk berubah, catat keputusan baru dan sesuaikan pengujian secara eksplisit.

## Titik awal

Audit awal memakai commit **`26dfb73`**, dengan pembanding scope performa **`b0ff0d2`**. Audit 4 November memakai **`d586eb6`**, sesudah penambahan keluarga refill samping dan perapatan kelompok. Semua revisi tersebut adalah referensi audit, bukan asumsi bahwa source mendatang tetap sama.

| Komponen | Keadaan saat penyusunan | Implikasi untuk optimasi |
| --- | --- | --- |
| `AAMediumPlanner.FillArea`, cache gap, `SortedSlots` | Jalur inti sama dengan pembanding scope lama | Temuan lama masih menjadi hipotesis yang layak diuji ulang. |
| `SortedSlots` | Sudah quicksort; timer mencakup transfer Collection ke array, sort, dan Collection hasil | Pisahkan biaya sebelum menyalahkan algoritme sort. |
| `OrderRefilledPattern` | Pengurutan hasil refill memakai aturan baris tersendiri | Jangan disamakan dengan urutan slot pencarian. |
| `AAQuantityFlow` dan `AAQuantityPlanner` | Ada target/carry/filler, optimasi quota final, serta blok/refill kelompok | Pisahkan waktu trial dari pekerjaan sesudah quota final. |
| Mark KissA dan refill | Anchor aktual/fallback, refill final dan per area memengaruhi hasil | Waktu serta jumlah lama tidak cukup sebagai baseline source sekarang. |
| FreeGapCache | Source tersedia; pengukuran lama belum menunjukkan manfaat waktu konsisten | Evaluasi A/B terpisah; kapasitas lebih besar belum tentu membantu. |
| Refill samping Medium | Kolom greedy 512 probe, pembanding baris 128 probe, lalu kontak rigid maksimal 32 pose kelompok | Pisahkan biaya dan kontribusi tiap keluarga dari seed yang sama; satu pose kelompok bukan satu probe objek. |

Rincian audit, angka historis, dan batas bukti berada pada topik AutoAlign tanggal 24 Oktober, 3 dan 4 November 2026 di [Development.log](../Development.log). Riwayat validasi pengguna tetap dicatat di sana.

## Acuan Case_07 setelah perubahan refill

Observasi log `AA-Debug-Case-07-Pg-3.txt` mulai baris **66677**, Medium/Custom, tiga pasangan dalam `/1` bebas, Otomatis, area **1000 × 1000 mm**, gap H/V **1 mm**. Hasil akhir **316 (105/105/106)**, bounds **998,948 × 987,584 mm**. Ini satu run acuan kualitas dan profil, belum seri benchmark berulang.

| Pekerjaan | Waktu teramati | Implikasi |
| --- | ---: | --- |
| Process total / planning | 549,828 / 542,814 s | Planning dominan pada kasus ini. |
| 11 trial kapasitas | 422,400 s | Sekitar 77,8% planning; prioritas utama investigasi. |
| 7 trial kapasitas yang tidak memenuhi vector target | 368,128 s | Sekitar 67,8% planning; hasil parsial bukan bukti batas kapasitas global. |
| Optimasi kuota final, fase full | 43,815 s | Terpisah dari 11 trial kapasitas. |
| Blok Minimum + refill | 26,465 s | Keluarga pembanding yang dipanggil setelah kuota final. |
| Refill samping keseluruhan | 44,492 s | Sekitar 8,2% planning; termasuk baris dan kontak, bukan tiga biaya independen. |
| Pembanding baris samping | 2,788 s | Kandidat 298 → 316, mengungguli kolom 311 sebanyak 5 objek. |
| Kontak rigid kelompok | 3,140 s | Jumlah tetap; 27/32 pose memperbaiki posisi 18 filler. |

Counter global gap **198,766 s**, overlap native **179,024 s**, contour capture **47,438 s**, proof pose **53,824 s**, dan slot sort **19,826 s** menunjukkan sasaran investigasi, bukan atribusi per trial. Waktu ini dapat saling mencakup atau berada di dalam tahap lain; jangan dijumlahkan untuk memperkirakan penghematan. Biaya greedy sendiri belum diukur terpisah dari setup, validasi dan pengurutan dalam tahap side.

Urutan perbaikan aktif: **O0 pengukuran → O4 trial/redundansi → O3 geometri/early rejection → O2 cache terisolasi sesuai bukti → O6 penjadwalan strategi**. O1 tetap tersedia bila profil baru menunjukkan biaya transfer/sort yang layak didahulukan. Pilih satu perubahan mekanis per eksperimen agar penyebab perbaikan atau regresinya dapat ditelusuri.

**Prasyarat correctness, 4 November — selesai pada scope Case_05:** deformasi free curve pada translasi node berurutan di `AAGeometry.AddPlaced` diperbaiki dengan transform seluruh kurva, binding identitas ke dokumen sebelum `TranslateBy`, serta proof pose/validator tetap. Pembanding partial baris tepi Straight Minimum juga tersedia. Penerimaan dan batas cakupan pengujian dicatat di Development.log. Scope berikutnya dimulai dari O0; tidak otomatis mengubah algoritme/budget optimasi. Angka Case_07 sebelum primitive translasi diperbaiki adalah referensi historis, bukan baseline performa aktif.

Retest berikutnya pada `AA-Debug-Geometry-Fix.txt` memuat lima Process selesai: Case_05 Minimum/Medium, Case_01 Minimum/Medium dan Case_07 Medium. Case_07 tetap 316 (105/105/106) dengan perapatan sama; observasi planning/Process **312,247/315,401 s**, trial kapasitas **221,399 s**, side inclusive **41,018 s**, capture **24/0,034 s** dan overlap **33296/150,656 s**. Ini calon acuan sesudah transform, belum seri A/B berulang; profil lama di atas bersifat historis. Pada versi sebelum partial Minimum, Straight Case_05 memperoleh 434 versus ZigZag 437. Status compile/penerimaan serta bukti runtime dicatat di Development.log, tanpa menganggap keberhasilan run sebagai penerimaan semua jalur.

Pembanding fase tepi Straight Minimum untuk satu Design non-rectangle bebas/satu area mempertahankan prioritas jumlah dan envelope: acuan Case_05 **437 dalam 304,5 × 461,5 mm**. Catatan penerimaan tetap di Development.log. Stage `Minimum partial edge rows` perlu dimasukkan dalam profil keluarga strategi. Run Case_01/Case_07 di Geometry-Fix mendahului patch partial terakhir; tetap bedakan observasi revisi tersebut dengan baseline berulang yang akan diambil sesudah instrumentasi O0.

**Pengujian primitive Case_12 tersedia:** harness sementara di project VBA yang sama menjalankan class `AAGeometry` asli pada delapan halaman, dua unit, variasi orientasi/penempatan dan translasi berulang, serta memeriksa proof/recapture dengan kontrol negatif. Source produksi dan state dokumen dipertahankan; module uji dibersihkan. Bukti, microbenchmark eksploratif dan batas cakupan dicatat di Development.log. Tahap ini tidak menggantikan Process end-to-end atau baseline berulang O0. Belum ada alasan dari hasil tersebut untuk mengubah validator atau budget; prioritas berikutnya tetap atribusi biaya trial/refill dan overlap pada workflow nyata.

## Prioritas awal yang dapat direvisi

ID dipakai untuk merujuk eksperimen, bukan menentukan urutan wajib.

| ID | Arah pekerjaan | Prioritas awal | Status | Bukti yang diperlukan untuk keputusan berikutnya |
| --- | --- | --- | --- | --- |
| O0 | Baseline terbaru, atribusi trial dan efisiensi strategi | Pertama | I1 trial tersedia di source, menunggu VBE/runtime; I2 kontribusi refill direncanakan | Vector trial, biaya/hasil strategi dan baseline berulang dapat dibandingkan. |
| O1 | Transfer data dan pengurutan slot | Setelah hotspot dominan, kecuali profil berubah | Direncanakan | Biaya transfer/sort diketahui; urutan identik; waktu bersih membaik. |
| O2 | Evaluasi cache pasangan pada pengisian bebas | Eksperimen terpisah | Cache sudah ada; manfaat belum ditetapkan | A/B cache aktif/nonaktif dengan source, profiling, dan input setara. |
| O3 | Early rejection, reuse proof/geometri, overlap native dan mark | Setelah atribusi O0; dapat mendahului O4 sesuai temuan | Investigasi terbuka | Predicate mahal yang dihemat, biaya cache/proof, invalidasi dan kesetaraan hasil. |
| O4 | Trial quota, kandidat redundan, blok dan filler/refill | Sasaran utama Case_07 sesudah O0 | Investigasi terbuka | Vector/context trial terpisah, pengulangan terbukti, hasil serta kualitas terjaga. |
| O5 | Pemeriksaan regresi lintas jalur yang terdampak | Mengikuti setiap perubahan | Berjalan per eksperimen | Bukti sesuai scope; tidak menunggu semua arah optimasi selesai. |
| O6 | Urutan strategi dan alokasi komputasi adaptif | Sesudah metrik stabil | Rencana; belum mengubah budget atau scheduler | Kontribusi marginal, efek urutan, kualitas, biaya tanpa debug dan kebijakan fallback diketahui. |

O0 menyediakan acuan untuk menilai manfaat, tetapi profil baru boleh memindahkan O3 atau O4 mendahului O1. O2 tidak menjadi syarat untuk mengerjakan O1. Eksperimen tanpa manfaat boleh dihentikan; keputusan dicatat tanpa menghapus riwayatnya.

### O0 — pengukuran trial dan kontribusi strategi

Langkah implementasi pertama dibatasi pada **instrumentasi**, tanpa mengubah urutan strategi, budget, cache keputusan atau validator. Gunakan `AADebugPerfSnapshot`/`AADebugPerfDelta` yang sudah ada pada batas tahap; hindari stopwatch atau output per edge/pasangan. Ringkasan tampil pada **`AADebugOn`**; **`AADebugOn True`** hanya jika investigasi geometri memerlukannya. Ukur overhead profiler dengan seri terpisah sebelum menganggap pengurangan waktu kecil sebagai manfaat algoritme.

Catat satu ringkasan per trial:

- ID attempt, fase capacity/full, vector `(Design, Quantity)` lengkap, target/aktual per Design, hasil fits/partial/rejected/cancelled, serta label kandidat terpilih.
- Konteks sumber/pairing/rotasi, seed/prefix dan versi state/cache dalam scope Process; target total sama tidak berarti trial yang sama.
- Durasi trial dan delta counter gap/overlap/weld/capture/proof/placement/slot; atribusi setup dan baseline Minimum tetap terlihat. Tambahkan hasil proof reuse/translation/recapture dan alasan penolakan pada level agregat jika diperlukan untuk memilih eksperimen.

Catat satu ringkasan per keluarga strategi yang benar-benar dipanggil (blok, refill lokal/relocate, side greedy, side rows kiri/kanan, kontak, partial edge atau independent-area sesuai jalur):

| Metrik | Makna |
| --- | --- |
| `seed`, `candidate`, `incumbent before/after` | Jumlah dan vector alokasi; bedakan baseline keluarga dengan hasil terbaik saat keluarga mulai. |
| Tambahan dari seed | `candidate - seed`, sesudah validasi keluarga; bukan jumlah percobaan yang pernah lolos sementara. |
| Peningkatan atas incumbent | `max(0, candidate - incumbent before)` bila kandidat diadopsi; perubahan ranking seri dicatat terpisah. |
| Kontribusi pada hasil akhir | Keluarga yang tetap dipilih dan jumlah akhir; kandidat yang kemudian tergantikan tidak dihitung sebagai tambahan permanen. |
| Waktu | Waktu keluarga termasuk persiapan/validasinya; parent/child diberi scope jelas. Setup bersama dicatat sekali, bukan dibebankan ulang pada semua keluarga. |
| Produktivitas | Tambahan dari seed/detik dan peningkatan atas incumbent/detik, masing-masing berlabel; waktu nol/terlalu pendek menjadi N/A, bukan infinity. |
| Biaya pencarian | Probe posisi, pose kelompok, kandidat valid/ditolak/duplikat, alasan bounds/mark/gap/overlap, cache hit/miss/bypass/replacement dan delta kernel. |
| Status | selected/replaced/tie/rejected/skipped/cancelled serta alasan batas; kegagalan atau Cancel tidak diberi produktivitas run sukses. |
| Kualitas tanpa tambahan | Perubahan compactness/bounds/alignment dan delta gerak; kontak rigid dinilai di sini, bukan dianggap tidak berguna karena menghasilkan nol objek. |

Contoh acuan: row menghasilkan **18/2,788 ≈ 6,46 objek/detik dari seed**, tetapi hanya **5/2,788 ≈ 1,79 objek/detik di atas incumbent kolom**. Kedua angka bukan kontribusi independen yang boleh dijumlahkan. Kolom menghasilkan kandidat +13 yang kemudian tergantikan, tetapi bukan berarti dapat dihapus untuk semua geometri. Kontak +0 memperbaiki posisi; durasinya tetap masuk biaya hasil terpilih. Produktivitas keseluruhan portfolio memakai hasil akhir dikurangi seed bersama dibagi seluruh waktu portfolio, termasuk keluarga yang kalah.

Perbandingan dari seed yang sama dan uji urutan/ablation terkontrol dipakai kemudian untuk mengatasi bias urutan. Jangan menjalankan semua ablation pada pemakaian normal atau menyimpan seluruh kurva kandidat hanya untuk statistik. Ringkasan scalar/array secukupnya dan lifetime satu Process menjadi pilihan awal.

#### Rencana implementasi instrumentasi O0

Rencana **4 November 2026**, audit batas pengukuran memakai source `8d3ea30`. **5 November: I1 diimplementasikan di source, menunggu compile VBE dan runtime pengguna; I2 tetap direncanakan.** Implementasi dibagi menjadi dua increment agar kebenaran log dan overhead dapat diperiksa sebelum memperluas cakupan. Tidak membutuhkan kontrol UserForm, registry, akses COM tambahan, atau module profiler baru; gunakan `AADebug.bas` yang sudah `Option Private Module` dan clock QPC/Timer yang sama.

| Increment | Batas source dan keluaran | Gate sebelum melanjutkan |
| --- | --- | --- |
| I1 — atribusi trial kuota | `AAQuantityPlanner.TryQuota`, konteks pemanggil `AAQuantityFlow`/allocator dan batas `AAMediumPlanner.Build`. Catat satu awal/akhir trial, vector, status, durasi dan delta counter global yang lifetime-nya cocok. | Trial dapat ditelusuri ke search/container tujuan; capacity/full, partial, error dan Cancel tidak tertukar. Hasil Process tetap sama. |
| I2 — atribusi strategi refill | Batas pemanggilan di `AAPlanner.ImproveMinimumGroup`/`ImproveMediumGroup`, lalu keluarga partial/final/independent di `AAMediumPlanner` sesuai jalur. Catat kandidat, incumbent, pilihan akhir, durasi dan work yang tersedia. | Alternatif dari seed sama tidak dijumlahkan sebagai gain independen; biaya parent/child dan budget bersama jelas; overhead terukur. |

Mulai implementasi dari I1. I2 boleh dipecah lagi menjadi portfolio kelompok, lalu partial/final/independent. Pengukuran setiap pattern/relocate internal, reason counter baru, proof-reuse/recapture rinci dan cache lookup merupakan perluasan terarah bila ringkasan awal belum menjawab hotspot; jangan langsung menambah counter pada semua loop geometri.

**Kontrak record trial:**

- ID run/search/attempt dan konteks container lokal Process. Nomor `/n` merupakan label tambahan, bukan identitas unik karena ada carry dan container tambahan. Catat jalur finite/free/modifier bila berlaku. Daftar sumber, pasangan dan pengaturan dicatat sekali per konteks, lalu dirujuk ID.
- Fase `capacity` atau `full`, target total dan vector requested lengkap menurut urutan Design. Ambil snapshot requested sebelum caller mengubah array kuota; total sama dengan vector berbeda tetap dua permintaan berbeda. Saat ada plan hasil, catat vector actual dari plan, bukan dari target.
- Outcome `fits`, `partial`, `rejected`, `cancelled` atau `error`. Partial menunjukkan pencarian ini belum memenuhi vector, bukan bukti kapasitas maksimum. Exception tanpa plan memberi actual `N/A`, bukan nol atau hasil trial sebelumnya. Bedakan outcome tersebut dari pemilihan kandidat akhir.
- Elapsed dan delta calls/seconds dari counter Process: capture, gap, overlap, weld, placement, proof pose, persiapan/sort slot yang sudah tersedia. Sertakan pekerjaan baseline Minimum di dalam trial dengan label child jika batasnya tersedia. Waktu tetap inclusive; belum menghitung exclusive time dari subtraction sembarang.
- Untuk cache, tulis scope/reset yang sebenarnya. Cache isi dapat bertahan antar trial, sementara statistik `mFill*`, `mGap*` dan `mReject*` di-reset pada awal `AAMediumPlanner.Build`. Nilai akhir Build dapat dipakai untuk Build tersebut; jangan mengurangkan snapshot lintas reset atau mencampurnya dengan hit/miss seumur search.

**Kontrak record refill:**

| Keluarga awal | Pengukuran dan batas penting |
| --- | --- |
| Blok Minimum + refill | Portfolio tersendiri sebelum refill samping; pattern/relocate di dalamnya masih inclusive pada increment awal. Catat seed serta kandidat yang lolos adaptasi/validasi ketika dipakai oleh Medium. |
| Side greedy | Durasi panggilan `GroupRefill(..., True)`, seed bersama, candidate dan incumbent; pisahkan dari timer side keseluruhan yang juga mencakup setup, rows, contact dan ordering. |
| Side rows kiri/kanan | Record masing-masing pemanggilan dengan seed yang sama. Budget probe 128 dipakai bersama kedua sisi; laporkan pemakaian per panggilan serta total portfolio, jangan menyatakan 128 per sisi. Setup forbidden/offset dicatat pada parent/setup, bukan dibebankan dua kali. |
| Kontak kelompok | Record hanya saat dipanggil pada kandidat row terpilih. Pose rigid kelompok maksimal 32 berbeda satuan dari probe penempatan objek. Count tetap; catat delta gerak dan bounds/ukuran celah yang sudah tersedia tanpa menghitung geometri baru. |
| Partial edge, final/free-template, independent area | Label jalur dan area/tepi; hitung gain terhadap input masing-masing. Kandidat partial yang ikut baseline bukan tambahan kedua kalinya pada gain akhir Medium. |

Setiap keluarga menyimpan `seed`, `candidate`, `best-before/after`, vector alokasi, raw gain dan gain atas incumbent. Catat diterima/tie/tidak membaik/ditolak serta alasan gate yang benar-benar tersedia. Pilihan keluarga akhir dicatat pada record portfolio, setelah perbandingan dan validasi; kandidat yang kalah tetap masuk biaya portfolio. Tidak menyimpan reference plan/Shape/Curve untuk kebutuhan statistik atau menambah field provenance pada plan produksi.

Satu instance keluarga dalam satu run belum cukup untuk mengatur budget adaptif. Agregasi analisis nantinya membedakan fase capacity/full/final, bentuk/kompleksitas sumber, seed dan alasan berhenti. Durasi terlalu pendek memakai `N/A` untuk rate; Cancel/error tetap memiliki elapsed tetapi dikeluarkan dari rate run sukses. Tahap kontak dinilai sebagai perbaikan posisi, bukan bersaing pada obj/detik dengan keluarga penambah objek.

**Lifetime, biaya dan jalur gagal:**

- `AADebugPerfSnapshot` adalah counter Process dan cocok untuk selisih pada batas trial. Sebaliknya, `AADebugSlotSnapshot` memakai counter yang di-reset oleh `AADebugCandidateStart` serta hanya aktif pada debug detail. Counter slot itu tidak boleh langsung dijadikan delta seluruh trial atau dianggap nol penolakan pada debug ringkas. Bila diperlukan, tambahkan agregasi numerik lifetime Process pada hook penolakan yang ada dalam increment terpisah.
- Debug nonaktif keluar sebelum serialisasi vector/alokasi record. Saat aktif, snapshot numerik diambil sebelum format/output; tulis ringkasan di batas tahap, tanpa log per edge/node/placement, pemeriksaan collision tambahan atau `DoEvents` tambahan. Interval UI 200 ms tetap.
- Scope yang bersarang memakai parent ID eksplisit; helper kecil dan array scalar cukup. Reset konteks pada awal Process dan bersihkan pada selesai/gagal/Cancel agar run berikutnya tidak mewarisi trial aktif. Semua status akhir harus berpasangan dengan awal yang tercatat; scope yang tidak pernah dipanggil tidak diberi waktu nol palsu.
- Simpan `Err.Number/Source/Description` sebelum logging; whitelist penolakan dan re-raise error/Cancel tetap persis. Kegagalan diagnostik tidak boleh menutupi error asli atau mengubah cleanup/Undo.

**Pengujian increment:**

1. Periksa helper ID/nesting/reset, vector requested/actual, zero-duration, partial/rejected dan Cancel/error; test adaptasi diberi label terpisah dari VBE/runtime. Cocokkan record dengan timer/counter lama tanpa menjumlahkan waktu nested.
2. Compile VBE, lalu Case_07 Page 3 dengan **`AADebugOn` ringkas**, input semula Custom 1000 × 1000/gap 1/Otomatis/tiga pasangan `/1` bebas. Ambil minimal dua run sebanding setelah instrumentasi; bandingkan count/vector, bounds, pose dan Z-order dengan acuan 316 (105/105/106), bukan memaksa angka tersebut jika source/input berbeda.
3. Satu kontrol Quantity finite/carry/`*`/`&` dan mixed source untuk I1 sesuai caller terdampak; Case_05 partial dan satu kontrol refill/pair untuk I2. Uji Cancel di tengah trial/refill lalu Process berikutnya, Undo/Redo, serta gap/mark/area pada jalur terkait. Pengukuran host tidak digantikan harness Case_12.
4. Ukur overhead sebagai seri terpisah pada revisi/input sama, debug ringkas dibanding debug nonaktif, dengan metode timing eksternal yang konsisten. Laporkan variasi/median/rentang; jangan menyebut selisih kecil sebagai speedup algoritme atau menetapkan persen overhead universal sebelum datanya ada.

Keluaran I1/I2 adalah profil yang bisa dipakai memilih eksperimen O4/O3: trial gagal mana yang mahal, predicate apa yang dominan, dan keluarga refill mana yang memberi manfaat dengan biaya berapa. Early rejection, reuse hasil geometri, deduplikasi dan scheduler tetap menunggu scope implementasi berikutnya.

#### Pelaksanaan baseline I1 — Case_07 Page 1 dan Page 3

Source yang diperbarui: Standard Module `AADebug`, Class Modules `AAQuantityPlanner`, `AAQuantityFlow`, `AAMediumPlanner` dan `AAProcessEngine`. Pastikan versi tersebut tersimpan/dimuat kembali setelah restart, lalu compile. I1 tidak mengubah kernel geometri, algoritme, budget atau refresh UI. Record trial hanya untuk `TryQuota`; template/direct build dapat tidak memiliki attempt. Atribusi strategi refill I2 belum termasuk pada baseline ini.

Rangkaian yang direncanakan pengguna, sedikitnya **12 Process** untuk dua halaman:

| Sesi | Kondisi Corel | Urutan minimum |
| --- | --- | --- |
| A | Instance awal, catat apakah sudah dipakai sebelumnya | P1-R1, P1-R2, P3-R1, P3-R2 |
| B | Restart sebelum sesi | P1-R1, P3-R1, P1-R2, P3-R2 |
| C | Restart sebelum sesi, khusus Page 1 | P1-R1, P1-R2 |
| D | Restart sebelum sesi, khusus Page 3 | P3-R1, P3-R2 |

Pertahankan source awal dan settings yang sama untuk halaman yang sama di seluruh sesi; Page 1 tidak dipaksa meniru setting Page 3. Catat label sesi/run di nama file atau catatan pengantar. Pada setiap restart aktifkan **`AADebugOn` ringkas** dan file logging kembali; mode detail bukan bagian seri ini. Pisahkan percobaan Cancel/input invalid dari run timing selesai. Run pertama setelah restart dibedakan dari run berikutnya; cache quota internal berumur search, sehingga selisih warm/cold host tidak otomatis berasal dari cache quota yang bertahan antarrun.

Laporkan semua nilai per run serta median/rentang per halaman dan kondisi sesi, jumlah/vector alokasi, bounds dan perbandingan visual. Jangan menggabungkan halaman atau kondisi warm/cold menjadi satu angka. Pengukuran overhead debug ringkas/nonaktif tetap seri terpisah; seri ini sendiri belum membuktikan speedup atau keuntungan setiap keluarga refill.

### O1 — pengurutan slot

Ukur atau bandingkan tiga pekerjaan: pembacaan Collection ke array, quicksort/pembandingan, dan pembuatan Collection hasil. Pertahankan slot, koordinat, urutan Y/X, tie, serta urutan penambahan tetangga. Jangan memakai pengurangan slot/probe sebagai jalan pintas untuk mempercepat sort.

Gunakan perbandingan hasil lengkap sebelum timing. Uji input kosong, satu slot, koordinat sama, banyak slot, dan pola yang sudah/hampir terurut sesuai kebutuhan perubahan. Pengukuran fixture hanya membantu memilih implementasi; manfaat Process tetap memerlukan run CorelDRAW.

### O2 — cache pasangan bebas

Pisahkan dari cache penolakan posisi dan cache kelompok `/n`. Ukur hit/miss/bypass/replacement, jumlah predicate yang benar-benar dihemat, biaya lookup/alokasi/cleanup, serta waktu total. Bandingkan cache aktif/nonaktif tanpa mencampurkan perubahan algoritme lain.

Pilihan hasilnya terbuka: pertahankan, batasi scope, ubah strategi, atau lepaskan pemakaian pada pengisian bebas. Belum ada keputusan untuk rollback source. Hasil satu pasangan tidak boleh dipakai sebagai hasil aman terhadap seluruh daftar tetangga yang terus bertambah.

### O3 — overlap dan mark

Pisahkan baseline Minimum, kandidat biasa, pola pasangan, centering/mark search, validasi, dan tahap akhir. Counter global tidak otomatis menunjukkan tahap yang dominan.

Jangan menghapus `WeldWith`/overlap hanya karena pemeriksaan gap lolos. Setiap reuse memerlukan kesetaraan geometri, posisi, toleransi, dan lifetime yang jelas. Anchor/fallback mark serta allowance harus tetap dibedakan.

Eksperimen setelah atribusi tersedia, masing-masing terpisah:

1. **Early rejection sebelum copy/native overlap.** Bounds dan broad phase sudah ada; ukur yang lolos serta biaya sebenarnya. `AAPlanner.CandidateFits` memeriksa overlap sebelum gap, sedangkan `AAMediumPlanner.Fits` memeriksa gap dahulu. Bandingkan biaya/rasio penolakan sebelum mempertimbangkan penyamaan urutan. Predicate gagal boleh menghentikan lebih awal, tetapi gap yang lolos tidak menghapus kewajiban overlap/containment, dan precheck sebelum copy tidak menggantikan validator geometri aktual.
2. **Reuse proof geometri yang immutable.** `AAPlacement.VerifyContourGeometry` sudah memoized per placement dan dibatalkan oleh Translate/penggantian collision; `CopyForPlan` menyiapkan proof baru. Ukur berapa proof berhasil sebagai translasi, recapture karena deformasi, serta pengulangan pada basis sama. Hanya reuse bukti dengan identitas/revisi geometri dan pose yang cocok; jangan menganggap copy atau translasi rencana selalu rigid pada Corel.
3. **Cache predicate pasangan, baru bila repeat rate memadai.** Bedakan cache precheck gap yang sudah ada dengan hasil gap/overlap pada copy aktual. Key harus mewakili kedua geometri terverifikasi, pose/transform, gap/toleransi dan predicate. Reuse berbasis offset relatif hanya setelah invariansi translasi aktual terbukti; beda koordinat absolut dapat memberi pembulatan host berbeda. Cache dibatasi ukuran/lifetime, tidak menyimpan error/Cancel sebagai jawaban, dan tetap punya fallback validator asli.
4. **Tetangga dan blocker.** Pertimbangkan index untuk jalur `AAPlanner.CandidateFits` yang masih memindai occupied bila profil membenarkannya; jalur Medium sudah memakai `AAOccupiedIndex`. Append-only, perpindahan lokal, penggantian prefix dan kontak kelompok memerlukan invalidasi berbeda. Last blocker dapat dicoba lebih dahulu hanya selama identitas/pose/konteksnya valid dan keseluruhan penerimaan tetap setara.

### O4 — quota, blok dan refill

Bedakan trial kapasitas, optimasi quota akhir, blok per Design, refill append-only, susun ulang filler, refill final, serta refill area independen. Gunakan jumlah/alokasi dan kerapian sebagai pembanding bersama waktu.

Prioritas pencarian berasal dari profil baru. Batas 144 kandidat/32768 probe filler dan 512 probe refill final yang ada bukan sasaran untuk diperkecil tanpa bukti kualitas. Ketentuan kelompok, sumber tetap, dan template Cut harus dipertahankan ketika pekerjaan berulang dikurangi.

Mulai dari biaya gagal Case_07, bukan menambah cache besar tanpa bukti pengulangan. Bedakan:

- Trial dengan total 298 tetapi vector berbeda, capacity/full, atau baseline/seed berbeda; bukan duplikat hanya karena label target sama.
- Kandidat posisi identik dengan geometri/orientasi/gap/context identik. Deduplikasi mempertahankan kemunculan pertama serta ranking/tie, tanpa pembulatan yang menggabungkan slot dekat tetapi berbeda.
- Penolakan terhadap blocker yang tetap pada occupied append-only; boleh dipakai ulang hanya ketika blocker dan konteks tetap. Hasil aman terhadap seluruh occupied tidak otomatis aman setelah daftar bertambah. Setelah objek dipindah/dihapus, bukti lama harus dibatalkan atau diverifikasi ulang.
- Snapshot sebelum centering versus rencana final setelah transform; jangan mengembalikan snapshot lama sebagai hasil aman tanpa validasi pose akhir.

Penghentian keluarga/trial memakai batas atas hanya jika batas itu benar-benar konservatif untuk seluruh kandidat legal dan vector target yang bersangkutan. Estimasi bbox, hasil trial gagal, hint kapasitas atau produktivitas rendah bukan bukti bahwa tidak ada susunan lebih baik. Perubahan pencarian yang menukar peluang jumlah dengan waktu dipisahkan dari optimasi mekanis.

### O6 — penjadwalan strategi adaptif

Tahapan ini menunggu O0 dan sekurangnya satu eksperimen mekanis terukur:

1. Jalankan kebijakan saat ini sambil mengumpulkan statistik, tanpa mengubah keputusan. Pisahkan konteks seperti sisa strip kanan/kiri, footprint/aspect ratio, kompleksitas kontur, rotasi legal, jumlah Design dan aturan Quantity; hindari model klasifikasi yang rumit sebelum data menunjukkan kebutuhan.
2. Uji **urutan lebih murah dahulu** pada replay/seri A/B, sambil tetap mencoba keluarga lain dengan budget semula. Rows yang murah pada Case_07 layak didahulukan sebagai eksperimen, bukan aturan universal. Jaga seed pembanding yang sama serta tie/ranking dan perapatan kandidat terpilih; sekadar mengubah urutan dapat mengubah hasil.
3. Bila urutan baru memberi incumbent lebih baik, hentikan keluarga lain hanya dengan batas atas legal yang tidak dapat mengalahkannya. Tanpa bukti tersebut, lanjutkan pencarian semula.
4. Time budget adaptif atau penghentian heuristik memerlukan scope produk dan kriteria penurunan kualitas yang disepakati tersendiri. Sisihkan kesempatan eksplorasi/fallback agar keluarga yang jarang dipilih tidak tersingkir selamanya. Pengukuran wall time saja dapat membuat hasil berbeda menurut mesin; sertakan work/probe budget dan uji repeatability.

Tidak ada perubahan otomatis berdasarkan satu run, statistik lintas sesi/registry, pemangkasan budget, atau fitur Maximum yang diaktifkan oleh roadmap ini.

## Baseline dan kasus uji awal

Matriks ini dapat berubah bila dokumen/sumber berubah atau kasus lain lebih mewakili penggunaan. ID adalah acuan dari log; pastikan sumber masih sesuai sebelum run.

| Kasus | Input awal | Peran |
| --- | --- | --- |
| Case_07 Page 3 | Design 5883/5877/5888 dengan Cut masing-masing, `/1` bebas; Custom 1000 × 1000, gap 1 mm, Otomatis | Sasaran pertama scope 4 November: trial mahal, portfolio refill, proof/capture dan overlap. Acuan kualitas 316 (105/105/106). |
| Case_06 Page 7 | Design ID 7980, standalone, Quantity kosong | Kurva kompleks; pembanding scope awal. |
| Case_06 Page 8 | Design ID 7935, standalone, Quantity kosong | Kurva kompleks lain; biaya gap/refinement/finalisasi. |
| Case_10 Page 2 | Standalone ID 88/86/90, masing-masing `/1` tanpa target angka | Trial quota dan blok/refill kelompok. |
| Case_06 Page 11 | Sumber seperti `AA-Debug-Case-06-Pg-11.txt` | Regresi refill final dan urutan hasil. |
| Kasus 03Nov | Rectangle ID 9/8, 70 × 50 mm, Quantity kosong tanpa `/n` | Kontrol refill beberapa area dibanding run tunggal. |

Mulai scope 4 November dari Case_07 Page 3 dengan instrumentasi yang tidak mengubah hasil, lalu minimal dua run sebanding. Case_06 Page 7/8 dan Case_10 Page 2 tetap menjadi pembanding lintas bentuk; Case_06 Page 11/03Nov menjadi regresi refill saat implementasi terkait berubah. Tambahkan kasus Cut bersama/masing-masing, mixed standalone, Quantity angka/modifier, preset 90/270, DieA dan Custom sesuai jalur yang terdampak; satu kasus tidak mewakili semuanya.

Pengaturan untuk kasus lain kecuali dicatat berbeda: **Medium/KissA, Otomatis, 320 × 470 mm, gap default**, serta debug/detail konsisten. **Case_07 tetap memakai Custom 1000 × 1000, gap H/V 1 mm**, bukan preset ini. Gap default tetap mengikuti tipe objek; rectangle tajam preset tidak dipaksa memakai gap kurva 0,5 mm. Simpan kondisi mark/layer karena anchor aktual memengaruhi zona forbidden; Case_07 Custom tanpa mark tidak membuktikan jalur KissA/DieA.

Nama log baseline baru yang diusulkan: `AA-Debug-Case-07-Pg-3-PerfBaseline.txt`, dengan pemisah run serta revisi/source yang jelas. Nama rencana lama `AA-Debug-03Nov-MediumBaseline.txt` tidak dianggap berisi bukti bila belum dibuat. File baru juga belum menjadi bukti sampai benar-benar dibuat dan diperiksa.

### Cara membandingkan

1. Catat commit atau hash file yang diubah, versi CorelDRAW, kasus/sumber, Quantity/pairing/rotasi, area/gap/mark, dan pengaturan debug. Cocokkan source yang dimuat di VBE dengan revisi yang dimaksud.
2. Gunakan sumber/geometri awal yang sama pada setiap run, bukan hasil duplikat run sebelumnya. Pulihkan dokumen dan catat perubahan lain yang dapat memengaruhi timing.
3. Mulai minimal dua run per kasus untuk melihat variasi. Jika hasil kecil atau tidak konsisten, tambah pengulangan dan bandingkan A/B bergantian; laporkan nilai per run serta median/rentangnya. Dua run saja belum membuktikan manfaat umum.
4. Rekam waktu Process dan planning, tahap relevan, counter panggilan/cache, jumlah per Design/area, bounds, serta perubahan susunan. Waktu inclusive tidak dijumlahkan seolah tahap independen.
5. Pisahkan pengukuran debug detail dari pengukuran penggunaan normal. Jika menguji tanpa debug, gunakan cara timing yang sama untuk kedua revisi dan catat sebagai seri terpisah.

Jumlah lama 112/19/87 adalah referensi historis, bukan batas atas atau baseline wajib source sekarang. Refill valid boleh meningkatkan hasil. Jangan menafsirkan waktu dari dua hasil berbeda sebagai benchmark performa yang sepenuhnya setara.

## Penerimaan dan status eksperimen

Perubahan diterima berdasarkan scope kasus yang diuji: hasil valid, kontrak terjaga, manfaat waktu bersih terukur, serta tidak ada regresi material pada jalur terdampak. Ambang manfaat praktis ditetapkan dari variasi baseline; tidak ada persentase wajib yang dibuat sebelum pengukuran.

- **Direncanakan:** hipotesis dan kasus pembanding sudah jelas; belum ada perubahan source.
- **Diimplementasikan, menunggu uji:** source tersedia; manfaat/native belum dibuktikan.
- **Diuji, belum diputuskan:** bukti tersedia tetapi belum cukup atau hasil campuran.
- **Diterima pada scope tertentu:** bukti dan batas kasus penerimaan dicatat.
- **Ditunda / dibatalkan:** alasan dan kemungkinan pemicu peninjauan ulang dicatat.

Pisahkan pemeriksaan statis, harness adaptasi/mock, compile VBE, runtime CorelDRAW, dan pemeriksaan visual/manual. Lulus salah satunya tidak otomatis berarti yang lain lulus. Catat validasi dan penerimaan pengguna hanya di [Development.log](../Development.log), lalu perbarui status ringkas di roadmap.

## Pemeliharaan roadmap

- Tinjau prioritas setelah baseline baru, perubahan planner/helper bersama, hasil eksperimen yang bertentangan, atau perubahan kebutuhan pengguna.
- Setiap perubahan arah memperbarui tanggal, status/ID terkait, dan alasan singkat. Tidak perlu mempertahankan urutan lama demi mengikuti dokumen.
- Simpan rencana aktif di sini, tindak lanjut singkat di [Idea.log](../Idea.log), serta diagnosis dan bukti di [Development.log](../Development.log). Hindari menyalin seluruh riwayat ke beberapa tempat.
- Sebelum melanjutkan setelah jeda, cocokkan source dan relevansi baseline. Status diterima pada revisi lama tidak membuktikan semua jalur baru aman.

| Tanggal | Perubahan arah |
| --- | --- |
| 2026-11-03 | Roadmap dibuat sebagai dokumen dinamis. Fokus awal Medium/helper bersama; prioritas pertama baseline source terbaru, lalu pilih eksperimen dari biaya yang terukur. |
| 2026-11-04 | Case_07 sesudah refill menjadi sasaran awal. O0 diperluas menjadi atribusi vector trial dan produktivitas mentah/marginal/final tiap strategi; O4/O3 didahulukan terhadap sort berdasarkan biaya trial. O6 menunda alokasi adaptif sampai data, efek urutan dan kontrak kualitas diketahui. Belum mengubah source algoritme atau budget. |
| 2026-11-04 | Prasyarat perbaikan primitive translasi dan partial Minimum ditutup pada scope teruji. Profil Case_07 sesudah transform dipakai sebagai acuan awal O0; baseline/instrumentasi baru tetap diperlukan sebelum klaim penghematan. |
| 2026-11-04 | Pengujian intensif primitive AddPlaced melalui class VBA asli pada Case_12 selesai. O0 tetap langkah berikutnya; microbenchmark dipisahkan dari pengukuran planner dan tidak memicu refactor primitive tanpa atribusi biaya. |
| 2026-11-04 | Rencana O0 dirinci menjadi I1 atribusi trial kuota dan I2 kontribusi/biaya refill. Kontrak record, lifetime/reset counter, nesting, Cancel/error dan pengujian overhead ditetapkan sebagai rencana; source VBA, budget dan scheduler belum diubah. |
| 2026-11-05 | I1 trial dipasang pada debug ringkas: ID konteks/attempt, vector requested/actual, outcome dan delta kerja; kernel/budget/scheduler tetap. Menunggu compile/runtime. Baseline pengguna direncanakan pada Case_07 P1/P3, sesi awal, restart bergantian dan restart per halaman; atribusi refill I2 tetap terpisah. |
