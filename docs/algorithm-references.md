# Referensi Algoritma AutoAlign

AutoAlign memakai heuristik nesting dua dimensi dengan batas pencarian. Kandidat mengikuti batas area, gap, zona mark, rotasi, Quantity, dan pasangan Cut Line. Alasan di bawah menjelaskan manfaat dan batas yang terlihat dari implementasi. Source VBA menjadi acuan untuk perilaku yang berlaku. [Panduan penggunaan dan batasan](../README.md#status-dan-batasan).

## Perencanaan susunan

### Pola Straight, ZigZag, dan refinemen Medium

Minimum memakai heuristik grid AA. Planner membandingkan Straight dengan ZigZag untuk bentuk yang memenuhi syarat. Jalur khusus rectangle memakai Straight. Partial ZigZag dapat menggeser baris tepi untuk membuka slot tambahan. Medium memakai Minimum sebagai baseline pada jalur yang sesuai, lalu mencoba orientasi, pola selang-seling dengan rotasi berlawanan, dan refinemen posisi. Bottom-left-fill dari Burke dan rekan memberi pembanding untuk keluarga heuristik packing dua dimensi. Pola dan urutan kandidat AA ditentukan oleh source proyek. [`BuildMinimum` dan `MinimumPartialEdges`](../src/classes/AAPlanner.cls), [`Build`, `BuildOppositeLattice` dan `FillArea`](../src/classes/AAMediumPlanner.cls), [Burke dkk. (2006)][burke].

Pola berulang memberi kandidat awal dengan biaya yang terbatasi. ZigZag bisa menambah objek atau memperlebar footprint, bergantung pada bentuknya. Rotasi dan refinemen menambah peluang mengisi celah sekaligus menambah pemeriksaan geometri. Hasil terbaik bergantung pada pola, sudut, dan posisi yang sempat diuji.

### Pengurutan slot dan pemilihan susunan

Medium memakai quicksort untuk mengurutkan slot menurut Y menurun, lalu X menaik. Quicksort menata banyak slot melalui partisi rekursif. Urutan AA memberi prioritas pada slot atas, lalu kiri dalam satu Y. Jika slot bersaing memakai ruang yang sama, urutan ini dapat memengaruhi hasil pengisian. [`SortSlots` dan `BeforeSlot`](../src/classes/AAMediumPlanner.cls), [Hoare (1962)][hoare].

Pembanding AA mengutamakan sedikitnya area, lalu banyaknya Design. Jika kedua nilai sama di Medium, kerapian pola dipertimbangkan sebelum luas footprint. Jika dimensi footprint dianggap setara, rotasi menjadi salah satu pemecah seri. Minimum memakai aturan tersendiri untuk seri Straight dan ZigZag. Prioritas ini mempertahankan tujuan produksi, dengan konsekuensi bahwa kandidat yang lebih rapi bisa terpilih meskipun footprint-nya lebih luas. [`BetterPlan` dan `BuildMinimum`](../src/classes/AAPlanner.cls).

### Trial Quantity dan refill saat Process

Trial Quantity memakai ulang template Cut Line bila syaratnya terpenuhi. Untuk jalur kuota, AA mempersempit rentang pencarian seperti pencarian biner untuk mencari jumlah yang lolos dalam satu area, lalu mencoba tambahan sesuai batas pembagian. Medium memakai estimasi awal dan hasil parsial untuk memilih trial berikutnya. Penyempitan rentang mengurangi jumlah trial mahal. Hasil trial heuristik tidak dijamin monoton, sehingga menurunkan batas atas setelah satu kegagalan bisa melewatkan kuota lebih tinggi yang mungkin lolos jika diuji. [`BuildGroup` dan `TryQuota`](../src/classes/AAQuantityPlanner.cls).

Refill adalah perbaikan lokal yang dikembangkan di AA. Planner mencoba slot tepi, area independen, atau jalur di samping blok sesuai bentuk dan mode, dengan tetap memeriksa gap, mark, area, dan alokasi. Meneruskan susunan yang sudah ada menghindari perencanaan ulang dari awal. Pencarian lokal bisa melewatkan celah yang baru terbuka setelah susunan dirombak lebih luas. [`RefillEdgeRows` dan `RefillIndependentAreas`](../src/classes/AAMediumPlanner.cls), [`ImproveMinimumGroup` dan `ImproveMediumGroup`](../src/classes/AAPlanner.cls).

### Optimize dan pergeseran terhadap mark

Optimize pada susunan yang sudah ada memakai heuristik pemadatan AA. Minimum merapatkan baris dan kolom terhadap anchor. Medium merapatkan kelompok dan objek individual, lalu mencoba memusatkan hasil jika lolos validasi. Translasi mempertahankan jumlah dan rotasi, sehingga ruang yang memerlukan perubahan orientasi berada di luar jangkauan pencarian ini. [AAOptimizeMinimum](../src/classes/AAOptimizeMinimum.cls), [AAOptimizeEngine](../src/classes/AAOptimizeEngine.cls).

Untuk mark, planner mencoba translasi kelompok dalam rentang yang dibatasi area. Translasi mempertahankan posisi relatif sehingga susunan dapat melewati mark tanpa membangun ulang pola. Pilihan tetap bergantung pada ruang gerak dan offset yang diuji. `AddPlaced` memakai matriks yang terikat ke unit dokumen untuk menggeser tiap salinan kurva sebagai satu kesatuan. Cara ini menjaga hubungan node dan handle dalam satu transformasi dan tetap memerlukan salinan kurva untuk operasi native. [`FindAreaShift`](../src/classes/AAPlanner.cls), [`AddPlaced`](../src/classes/AAGeometry.cls).

Pada pencarian Custom dengan beberapa zona, `CustomShiftMarksFit` memeriksa lebih dulu footprint yang terakhir membentur mark, memakai offset probe saat ini. Benturan yang berulang memungkinkan probe langsung ditolak. Jika footprint itu lolos, pemeriksaan berlanjut ke footprint lain. Penghematan bergantung pada seberapa sering penghalang yang sama muncul pada probe berikutnya. [`CustomShiftMarksFit`](../src/classes/AAPlanner.cls).

## Pemeriksaan geometri

### Subdivisi Bézier kubik dengan de Casteljau

Algoritma de Casteljau membagi kurva kubik melalui interpolasi berulang atas empat titik kontrol. AA membagi pada parameter `t = 0,5`. `Flatten` memakai jarak terbesar dari kedua titik kontrol interior ke segmen penghubung ujung sebagai batas galat berbasis convex hull. Cabang berhenti saat galat memenuhi ambang flatness atau batas kedalaman tercapai. Batas galat tiap segmen tetap disimpan dan diperhitungkan dalam pemeriksaan gap. [`Flatten` dan `ClearOf`](../src/classes/AAContour.cls), [Farin (2002), bab 3 dan §4.6][farin].

Subdivisi adaptif memakai sedikit segmen pada bagian yang hampir lurus dan menambah segmen di bagian yang membutuhkan aproksimasi lebih rapat. Flatness yang lebih ketat dapat menambah biaya pembuatan kontur, memori, dan pasangan segmen yang perlu diperiksa. Batas kedalaman juga berarti sebagian segmen bisa menyimpan galat di atas ambang flatness.

### Penyaringan dengan bounding box

AA memakai bounding box sejajar sumbu atau AABB sebagai pemeriksaan awal (*broad phase*). `BoundsClearOf` dapat melewati pemeriksaan jarak segmen jika jarak minimum antarkotak kurva asli sudah memenuhi gap dalam toleransi. Bounding box mudah dibandingkan. Kotak yang berdekatan atau bertumpuk masih menyisakan kemungkinan bahwa kontur aslinya berjauhan, sehingga jalur detail tetap diperlukan pada kasus tersebut. [`BoundsClearOf` dan `ClearOf`](../src/classes/AAContour.cls), [Le Grand (2007)][legrand].

### Grid spasial seragam

Grid seragam membagi ruang menjadi sel berukuran tetap. AA memakai satu indeks untuk menyaring segmen kontur dan satu indeks lain untuk mencari objek yang sudah ditempatkan di sekitar kandidat. Query memperhitungkan margin yang diperlukan dan meneruskan kandidat ke pemeriksaan geometri. Sel padat menghasilkan lebih banyak kandidat, sementara objek panjang atau query lebar dapat menambah biaya penelusuran sel. Jalur fallback menjaga cakupan saat batas indeks tercapai. [`BuildEdgeIndex` dan `CandidateEdgeIndexes`](../src/classes/AAContour.cls), [`Nearby`](../src/classes/AAOccupiedIndex.cls), [Franklin (1984)][franklin].

### Overlap native dan cache pemeriksaan

AA memakai identitas luas irisan `luas(A) + luas(B) − luas(A ∪ B)` untuk kurva tertutup. `WeldWith` menghasilkan kurva gabungan yang luasnya dipakai dalam tes. Kurva terbuka memakai pemeriksaan persilangan atau apakah lintasan masuk ke kurva tertutup. Operasi native memeriksa kurva CorelDRAW secara langsung, dengan biaya panggilan dan toleransi luas untuk jalur kurva tertutup. [`CurvesOverlap`](../src/classes/AAGeometry.cls), [CorelDRAW SDK `Curve.WeldWith`](https://community.coreldraw.com/sdk/api/draw/25/m/curve.weldwith).

Cache AA memakai ulang hasil gap satu pasangan ketika identitas dan orientasi kontur, posisi, serta gap efektif cocok. Cache penolakan untuk satu posisi terikat pada daftar penghalang yang hanya bertambah selama fase terkait. Penggunaan ulang mengurangi pemeriksaan saat trial mengunjungi posisi yang sama. Cache memerlukan memori dan biaya pencarian key, sehingga manfaatnya bergantung pada frekuensi pengulangan. [`QuotaClearanceAt` dan `BeginQuotaRejectCache`](../src/classes/AAMediumPlanner.cls).

## Referensi publik

**Dipakai** menandai dokumentasi untuk API yang dipanggil AA. **Pembanding** menandai sumber matematika atau literatur yang menjelaskan keluarga teknik terkait. **Bacaan lanjutan** menandai pendekatan untuk dipelajari. Atribusi atas pilihan desain AA mengikuti bukti pada source dan riwayat proyek.

- **Dipakai** · CorelDRAW 2024 SDK. `Shape.DisplayCurve` membaca kontur sumber, `Curve.WeldWith` mendukung pemeriksaan overlap kurva tertutup, dan `ApplyTransformMatrix` menerapkan transformasi pada salinan kurva. [Shape.DisplayCurve](https://community.coreldraw.com/sdk/api/draw/25/p/shape.displaycurve), [Curve.WeldWith](https://community.coreldraw.com/sdk/api/draw/25/m/curve.weldwith), [Curve](https://community.coreldraw.com/sdk/api/draw/25/c/curve).
- **Dipakai** · Microsoft Learn. `Scripting.Dictionary` menyediakan penyimpanan berbasis key untuk indeks spasial. `QueryPerformanceCounter` dipakai untuk mengukur durasi pada debug AA. [Dictionary object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object), [QueryPerformanceCounter](https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter), [AADebug](../src/modules/AADebug.bas).
- **Pembanding** · Gerald Farin (2002). *Curves and Surfaces for CAGD*, edisi ke-5, bab 3 dan §4.6. Dasar de Casteljau dan subdivisi kurva Bézier. [Daftar isi dari penulis][farin].
- **Pembanding** · W. Randolph Franklin (1984). *Adaptive Grids for Geometric Operations*. Cartographica, 21(2–3), 161–167. Grid satu tingkat untuk menyaring kandidat pemeriksaan spasial. [Makalah dari penulis][franklin].
- **Pembanding** · Scott Le Grand (2007). *Broad-Phase Collision Detection with CUDA*. GPU Gems 3, bab 32. Menjelaskan pemisahan pemeriksaan awal berbasis bounding volume dan pemeriksaan detail. [NVIDIA Developer][legrand].
- **Pembanding** · Edmund Burke, Robert Hellier, Graham Kendall, dan Glenn Whitwell (2006). *A New Bottom-Left-Fill Heuristic Algorithm for the Two-Dimensional Irregular Packing Problem*. Operations Research, 54(3), 587–601. Pembanding untuk pengisian celah pada packing bentuk tak beraturan. [Makalah dari penulis][burke].
- **Pembanding** · C. A. R. Hoare (1962). *Quicksort*. The Computer Journal, 5(1), 10–16. Referensi untuk pengurutan dengan partisi rekursif. [Oxford Academic][hoare].
- **Bacaan lanjutan** · Julia A. Bennell dan Xiang Song (2008). *A comprehensive and robust procedure for obtaining the nofit polygon using Minkowski sums*. Computers & Operations Research, 35(1), 267–281. Membahas pembentukan no-fit polygon untuk merepresentasikan posisi relatif dalam nesting bentuk tak beraturan. [University of Southampton](https://eprints.soton.ac.uk/154799/).

[farin]: https://www.farinhansford.com/books/cagd/toc.html
[franklin]: https://wrfranklin.org/wiki/Research/p/23-cartographica86-grid.pdf
[legrand]: https://developer.nvidia.com/gpugems/gpugems3/part-v-physics-simulation/chapter-32-broad-phase-collision-detection-cuda
[burke]: https://www.graham-kendall.com/papers/bhkw2006.pdf
[hoare]: https://academic.oup.com/comjnl/article-abstract/5/1/10/395338

## Dokumen dan bukti pendukung

### Tersedia di Git

- [Roadmap Optimasi](../ROADMAP-Optimization.md) mencatat baseline performa, trial kuota, dan pengukuran keluarga strategi refill.
- [Roadmap Maximum](../ROADMAP-Maximum.md) mencatat rencana pencarian dan pemilihan kandidat untuk mode Maximum.
- [Roadmap Optimize](../ROADMAP-Optimize.md) mencatat rencana pemadatan dan pengisian celah pada container yang sudah ada.

### Lokal — khusus komputer pengembang

Path berikut dihitung relatif dari folder `AutoAlign/docs`. Target berada di luar repo AutoAlign atau di folder yang diabaikan Git. Akses bergantung pada salinan artefak di komputer pengembang.

- Riwayat keputusan, investigasi, dan validasi yang dilaporkan pengguna. `../../Development.log`.
- Inspeksi baseline pose lookup. `../tests/pose-lookup-run/Inspection.md`.
- Log Case_07 untuk evaluasi indeks segmen dengan mark. `../tests/edge-row-run/case-07-page-3-with-mark.txt`.
- Log penolakan awal terhadap mark pada Minimum dan Medium. `../tests/mark-blocker-run/case-07-page-3.txt.txt` dan `../tests/mark-blocker-run/case-07-page-3-medium.txt`.
- Pengujian adaptasi untuk indeks segmen dan aturan pergeseran. Bukti Compile VBE, run native, serta pemeriksaan manual dicatat di Development.log. `../tests/verify_edge_row_index.py` dan `../tests/verify_mark_shift_allowance.py`.
