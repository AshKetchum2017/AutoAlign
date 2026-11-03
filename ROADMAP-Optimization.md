# Roadmap Optimasi AutoAlign

Dibuat dan diperbarui: **3 November 2026**.

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

Audit source terakhir memakai commit **`26dfb73`**. Pembanding scope performa sebelumnya adalah **`b0ff0d2`**. Keduanya adalah referensi audit, bukan asumsi bahwa source mendatang tetap sama.

| Komponen | Keadaan saat penyusunan | Implikasi untuk optimasi |
| --- | --- | --- |
| `AAMediumPlanner.FillArea`, cache gap, `SortedSlots` | Jalur inti sama dengan pembanding scope lama | Temuan lama masih menjadi hipotesis yang layak diuji ulang. |
| `SortedSlots` | Sudah quicksort; timer mencakup transfer Collection ke array, sort, dan Collection hasil | Pisahkan biaya sebelum menyalahkan algoritme sort. |
| `OrderRefilledPattern` | Pengurutan hasil refill memakai aturan baris tersendiri | Jangan disamakan dengan urutan slot pencarian. |
| `AAQuantityFlow` dan `AAQuantityPlanner` | Ada target/carry/filler, optimasi quota final, serta blok/refill kelompok | Pisahkan waktu trial dari pekerjaan sesudah quota final. |
| Mark KissA dan refill | Anchor aktual/fallback, refill final dan per area memengaruhi hasil | Waktu serta jumlah lama tidak cukup sebagai baseline source sekarang. |
| FreeGapCache | Source tersedia; pengukuran lama belum menunjukkan manfaat waktu konsisten | Evaluasi A/B terpisah; kapasitas lebih besar belum tentu membantu. |

Rincian audit, angka historis, dan batas bukti berada pada topik AutoAlign tanggal 24 Oktober dan 3 November 2026 di [Development.log](../Development.log). Riwayat validasi pengguna tetap dicatat di sana.

## Prioritas awal yang dapat direvisi

ID dipakai untuk merujuk eksperimen, bukan menentukan urutan wajib.

| ID | Arah pekerjaan | Prioritas awal | Status | Bukti yang diperlukan untuk keputusan berikutnya |
| --- | --- | --- | --- | --- |
| O0 | Baseline source terbaru dan pemetaan waktu per jalur | Pertama | Audit selesai; pengukuran ulang menunggu | Input/revisi tercatat, hasil dan waktu berulang dapat dibandingkan. |
| O1 | Transfer data dan pengurutan slot | Kandidat awal | Direncanakan | Biaya transfer/sort diketahui; urutan identik; waktu bersih membaik. |
| O2 | Evaluasi cache pasangan pada pengisian bebas | Eksperimen terpisah | Cache sudah ada; manfaat belum ditetapkan | A/B cache aktif/nonaktif dengan source, profiling, dan input setara. |
| O3 | Biaya overlap native dan pencarian mark | Berdasarkan profil | Investigasi terbuka | Lokasi pengulangan dominan serta alasan reuse/pengurangan kerja yang tetap benar. |
| O4 | Trial quota, kandidat blok, dan filler/refill | Berdasarkan profil | Investigasi terbuka | Biaya tiap tahap terpisah; alokasi, jumlah dan kualitas hasil terjaga. |
| O5 | Pemeriksaan regresi lintas jalur yang terdampak | Mengikuti setiap perubahan | Berjalan per eksperimen | Bukti sesuai scope; tidak menunggu semua arah optimasi selesai. |

O0 menyediakan acuan untuk menilai manfaat, tetapi profil baru boleh memindahkan O3 atau O4 mendahului O1. O2 tidak menjadi syarat untuk mengerjakan O1. Eksperimen tanpa manfaat boleh dihentikan; keputusan dicatat tanpa menghapus riwayatnya.

### O1 — pengurutan slot

Ukur atau bandingkan tiga pekerjaan: pembacaan Collection ke array, quicksort/pembandingan, dan pembuatan Collection hasil. Pertahankan slot, koordinat, urutan Y/X, tie, serta urutan penambahan tetangga. Jangan memakai pengurangan slot/probe sebagai jalan pintas untuk mempercepat sort.

Gunakan perbandingan hasil lengkap sebelum timing. Uji input kosong, satu slot, koordinat sama, banyak slot, dan pola yang sudah/hampir terurut sesuai kebutuhan perubahan. Pengukuran fixture hanya membantu memilih implementasi; manfaat Process tetap memerlukan run CorelDRAW.

### O2 — cache pasangan bebas

Pisahkan dari cache penolakan posisi dan cache kelompok `/n`. Ukur hit/miss/bypass/replacement, jumlah predicate yang benar-benar dihemat, biaya lookup/alokasi/cleanup, serta waktu total. Bandingkan cache aktif/nonaktif tanpa mencampurkan perubahan algoritme lain.

Pilihan hasilnya terbuka: pertahankan, batasi scope, ubah strategi, atau lepaskan pemakaian pada pengisian bebas. Belum ada keputusan untuk rollback source. Hasil satu pasangan tidak boleh dipakai sebagai hasil aman terhadap seluruh daftar tetangga yang terus bertambah.

### O3 — overlap dan mark

Pisahkan baseline Minimum, kandidat biasa, pola pasangan, centering/mark search, validasi, dan tahap akhir. Counter global tidak otomatis menunjukkan tahap yang dominan.

Jangan menghapus `WeldWith`/overlap hanya karena pemeriksaan gap lolos. Setiap reuse memerlukan kesetaraan geometri, posisi, toleransi, dan lifetime yang jelas. Anchor/fallback mark serta allowance harus tetap dibedakan.

### O4 — quota, blok dan refill

Bedakan trial kapasitas, optimasi quota akhir, blok per Design, refill append-only, susun ulang filler, refill final, serta refill area independen. Gunakan jumlah/alokasi dan kerapian sebagai pembanding bersama waktu.

Prioritas pencarian berasal dari profil baru. Batas 144 kandidat/32768 probe filler dan 512 probe refill final yang ada bukan sasaran untuk diperkecil tanpa bukti kualitas. Ketentuan kelompok, sumber tetap, dan template Cut harus dipertahankan ketika pekerjaan berulang dikurangi.

## Baseline dan kasus uji awal

Matriks ini dapat berubah bila dokumen/sumber berubah atau kasus lain lebih mewakili penggunaan. ID adalah acuan dari log; pastikan sumber masih sesuai sebelum run.

| Kasus | Input awal | Peran |
| --- | --- | --- |
| Case_06 Page 7 | Design ID 7980, standalone, Quantity kosong | Kurva kompleks; pembanding scope awal. |
| Case_06 Page 8 | Design ID 7935, standalone, Quantity kosong | Kurva kompleks lain; biaya gap/refinement/finalisasi. |
| Case_10 Page 2 | Standalone ID 88/86/90, masing-masing `/1` tanpa target angka | Trial quota dan blok/refill kelompok. |
| Case_06 Page 11 | Sumber seperti `AA-Debug-Case-06-Pg-11.txt` | Regresi refill final dan urutan hasil. |
| Kasus 03Nov | Rectangle ID 9/8, 70 × 50 mm, Quantity kosong tanpa `/n` | Kontrol refill beberapa area dibanding run tunggal. |

Mulai dengan tiga baris pertama. Dua baris berikut menjadi regresi refill saat implementasi terkait berubah. Tambahkan kasus Cut bersama/masing-masing, mixed standalone, Quantity angka/modifier, preset 90/270, DieA dan Custom sesuai jalur yang terdampak; tidak menganggap tiga kasus awal mewakili semuanya.

Pengaturan awal: **Medium/KissA, Otomatis, 320 × 470 mm, gap default**, serta debug/detail konsisten. Gap default tetap mengikuti tipe objek; rectangle tajam preset tidak dipaksa memakai gap kurva 0,5 mm. Simpan kondisi mark/layer karena anchor aktual memengaruhi zona forbidden.

Nama log baseline yang direncanakan: `AA-Debug-03Nov-MediumBaseline.txt`. Bila tanggal atau scope pengukuran berubah, gunakan nama baru yang jelas dan catat pemetaannya; file tersebut belum menjadi bukti sampai benar-benar dibuat dan diperiksa.

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
