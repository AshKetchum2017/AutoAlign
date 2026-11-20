# Referensi Algoritma AutoAlign

AutoAlign menggunakan heuristik nesting dua dimensi dengan batas pencarian. Planner membandingkan kandidat yang memenuhi batas area, gap, zona mark, rotasi, serta aturan Quantity dan pasangan Cut Line. Source di bawah menjadi acuan untuk perilaku implementasi. [Panduan penggunaan dan batasan](../README.md#status-dan-batasan).

## Algoritma yang dipakai

| Bagian | Pendekatan | Source |
| --- | --- | --- |
| Minimum | Membandingkan susunan Straight dan ZigZag untuk bentuk yang memenuhi syarat. Jalur khusus rectangle memakai Straight. Untuk kandidat yang cocok, partial ZigZag mencoba menggeser baris tepi agar slot tambahan muat. | [`BuildMinimum`, `MinimumPartialEdges` dan `ImproveMinimumGroup`](../src/classes/AAPlanner.cls) |
| Medium | Memakai Minimum sebagai baseline pada jalur yang sesuai. Mencoba variasi orientasi, pola selang-seling dengan rotasi berlawanan bila bentuknya mendukung, dan penyempurnaan slot. Quicksort mengurutkan slot mengikuti Y menurun, lalu X menaik. | [`Build`, `BuildOppositeLattice`, `FillArea` dan `SortSlots`](../src/classes/AAMediumPlanner.cls) |
| Pemilihan susunan | Pembanding umum mengutamakan sedikitnya area, lalu banyaknya Design. Jika jumlah area dan Design sama di Medium, kerapian pola dipertimbangkan sebelum luas footprint. Jika dimensi footprint dianggap setara, pembanding melihat rotasi. Minimum memiliki aturan pemecah seri sendiri untuk kandidat Straight dan ZigZag. | [`BetterPlan` dan `BuildMinimum`](../src/classes/AAPlanner.cls) |
| Trial Quantity | Memakai ulang template Cut Line bila syaratnya terpenuhi. Untuk kelompok dengan bentuk berbeda, planner mencari kuota yang lolos dalam satu area dengan mempersempit rentang pencarian, lalu mencoba tambahan sesuai batas pembagian. Medium memakai estimasi awal dan hasil trial parsial untuk memilih kuota yang diuji berikutnya. | [`BuildGroup` dan `TryQuota`](../src/classes/AAQuantityPlanner.cls) |
| Refill saat Process | Mencoba mengisi slot tepi, area independen, atau jalur di samping blok. Strategi dipilih sesuai bentuk, mode, dan kelompok. Kandidat tambahan tetap mengikuti pemeriksaan gap, mark, batas area, dan alokasi. | [Refill Medium](../src/classes/AAMediumPlanner.cls), [`ImproveMinimumGroup` dan `ImproveMediumGroup`](../src/classes/AAPlanner.cls) |
| Kontur dan collision | Membagi Bézier kubik secara adaptif di titik tengah dengan metode de Casteljau. Setiap segmen menyimpan estimasi galat terhadap kurva untuk pemeriksaan gap. Pemeriksaan overlap memakai operasi kurva native CorelDRAW sesuai jenis lintasan. | [`Flatten` dan `ClearOf`](../src/classes/AAContour.cls), [`CurvesOverlap`](../src/classes/AAGeometry.cls) |
| Penolakan awal dan cache | Bounding box dan grid spasial menyaring segmen maupun objek sekitar. Pada jalur yang mendukungnya, cache mengurangi pemeriksaan berulang dengan orientasi, posisi, dan gap yang sama. Kandidat yang tersisa menjalani pemeriksaan geometri sesuai jalurnya. | [`BuildEdgeIndex` dan `ClearOf`](../src/classes/AAContour.cls), [`Nearby`](../src/classes/AAOccupiedIndex.cls), [`QuotaClearanceAt`](../src/classes/AAMediumPlanner.cls) |
| Pergeseran terhadap mark | Mencoba pergeseran kelompok secara terbatas dan memeriksa footprint terhadap area serta zona mark. Translasi kurva bekerja pada salinan dengan matriks yang terikat ke unit dokumen. | [`FindAreaShift`](../src/classes/AAPlanner.cls), [`AddPlaced`](../src/classes/AAGeometry.cls) |
| Optimize susunan | Minimum merapatkan baris dan kolom terhadap anchor. Medium merapatkan kelompok dan objek individual, kemudian mencoba memusatkan susunan yang lolos validasi. Keduanya mempertahankan jumlah dan rotasi. | [AAOptimizeMinimum](../src/classes/AAOptimizeMinimum.cls), [AAOptimizeEngine](../src/classes/AAOptimizeEngine.cls) |

## Referensi publik

**Dipakai** menandai dokumentasi API yang dipanggil oleh source. **Pembanding** menandai literatur untuk menjelaskan teknik yang serupa. **Bacaan lanjutan** menandai pendekatan yang dapat dipelajari untuk nesting berikutnya. Status literatur mengikuti bukti penggunaan yang tersedia di proyek.

- **Dipakai** — CorelDRAW 2024 SDK. `Shape.DisplayCurve` membaca kontur sumber, `Curve.WeldWith` mendukung pemeriksaan overlap kurva tertutup, dan `ApplyTransformMatrix` menerapkan transformasi pada salinan kurva. [Shape.DisplayCurve](https://community.coreldraw.com/sdk/api/draw/25/p/shape.displaycurve), [Curve.WeldWith](https://community.coreldraw.com/sdk/api/draw/25/m/curve.weldwith), [Curve](https://community.coreldraw.com/sdk/api/draw/25/c/curve).
- **Dipakai** — Microsoft Learn. `Scripting.Dictionary` menyediakan penyimpanan berbasis key untuk indeks spasial, sementara `QueryPerformanceCounter` dipakai untuk mengukur durasi pada debug AA. [Dictionary object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object), [QueryPerformanceCounter](https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter), [AADebug](../src/modules/AADebug.bas).
- **Pembanding** — C. A. R. Hoare (1962). *Quicksort*. The Computer Journal, 5(1), 10–16. Menjelaskan keluarga algoritma pengurutan yang juga dipakai untuk slot Medium. Urutan Y/X ditetapkan oleh `BeforeSlot` di source AA. [Artikel Oxford Academic](https://academic.oup.com/comjnl/article-abstract/5/1/10/395338).
- **Bacaan lanjutan** — Julia A. Bennell dan Xiang Song (2008). *A comprehensive and robust procedure for obtaining the nofit polygon using Minkowski sums*. Computers & Operations Research, 35(1), 267–281. Membahas pembentukan no-fit polygon untuk merepresentasikan posisi relatif dalam nesting bentuk tak beraturan. [Makalah di University of Southampton](https://eprints.soton.ac.uk/154799/).

## Dokumen dan bukti pendukung

### Tersedia di Git

- [Roadmap Optimasi](../ROADMAP-Optimization.md) mencatat baseline performa, trial kuota, dan pengukuran keluarga strategi refill.
- [Roadmap Maximum](../ROADMAP-Maximum.md) mencatat rencana pencarian dan pemilihan kandidat untuk mode Maximum.
- [Roadmap Optimize](../ROADMAP-Optimize.md) mencatat rencana pemadatan dan pengisian celah pada container yang sudah ada.

### Lokal — khusus komputer pengembang

Tautan berikut relatif terhadap workspace pengembang. Target berada di luar repo AutoAlign atau di folder yang diabaikan Git. Aksesnya bergantung pada salinan artefak di komputer pengembang.

- Riwayat keputusan, investigasi, dan validasi yang dilaporkan pengguna. [Development.log](../../Development.log).
- Ringkasan baseline untuk pemeriksaan penempatan pada trial kuota dan log Case_07 untuk evaluasi indeks segmen serta penolakan awal terhadap mark. [Inspeksi pose lookup](../tests/pose-lookup-run/Inspection.md), [Case_07 dengan mark](../tests/edge-row-run/case-07-page-3-with-mark.txt), [Case_07 Minimum](../tests/mark-blocker-run/case-07-page-3.txt.txt), [Case_07 Medium](../tests/mark-blocker-run/case-07-page-3-medium.txt).
- Pengujian adaptasi untuk indeks segmen dan aturan pergeseran. Compile VBE, run native, dan pemeriksaan manual memiliki bukti masing-masing di Development.log. [Uji indeks segmen](../tests/verify_edge_row_index.py), [Uji pergeseran terhadap mark](../tests/verify_mark_shift_allowance.py).
