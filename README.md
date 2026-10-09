# AutoAlign

AutoAlign adalah macro VBA untuk CorelDRAW yang dirancang untuk menyusun dan menggandakan objek **Design** beserta **Cut Line** secara otomatis di dalam area produksi.

Macro ini membantu mengurangi pekerjaan manual saat menentukan jumlah objek yang muat, mengatur jarak, mencoba pola susunan, dan menjaga hasil agar tetap berada di dalam area serta menghindari zona mark.

AutoAlign menyediakan mode **Minimum** dan **Medium**, model area **KissA**, **DieA**, dan **Custom**, serta pengaturan **Quantity** untuk jumlah tertentu, pengisian sisa area, atau pembagian beberapa Design dalam satu container.

## Main Features

### Object Container

Objek yang akan diproses didaftarkan ke `lbxObjects` sebagai:

- **Design** melalui `cmdSetDesign`.
- **Cut Line** melalui `cmdSetCutLine`.

Daftar menyimpan peran, `StaticID`, dan Quantity setiap Design. Objek dengan ID yang sama tidak ditambahkan dua kali.

Pendaftaran beberapa objek sekaligus mengikuti urutan baris dari atas ke bawah, lalu kiri ke kanan. Baris ditentukan dari pusat Y tertinggi dengan toleransi **1 mm** terhadap objek pertama pada baris tersebut.

`lbxObjects` mendukung pilihan banyak baris menggunakan **Ctrl/Shift**. `cmdRemove` menghapus baris terpilih dan `cmdClear` mengosongkan daftar; keduanya mengubah daftar tanpa menghapus objek di dokumen.

Semua objek dalam daftar harus berasal dari **satu dokumen dan satu halaman**, berada langsung pada halaman atau berupa group induk, serta berada pada layer yang dapat diedit.

### Design dan Cut Line

AutoAlign mendukung tiga susunan input:

| Input | Perilaku |
| --- | --- |
| Design tanpa Cut Line | Susunan memakai siluet luar Design sebagai acuan. |
| Satu Cut Line | Cut Line digunakan bersama oleh Design yang didaftarkan. |
| Cut Line sebanyak Design | Design ke-1 dipasangkan dengan Cut Line ke-1, dan seterusnya menurut urutan masing-masing peran dalam daftar. |

Untuk pasangan, **Cut Line menjadi acuan pola, gap, dan collision antarunit**. Bounds Design tetap diperiksa agar muat di dalam area. Artwork Design berpasangan dapat bertumpuk selama acuan Cut Line memenuhi aturan susunan.

Untuk pasangan selain rectangle bersudut runcing, validasi ukuran mengizinkan Design lebih kecil maksimal **0,2 mm** atau lebih besar maksimal **2 mm** dari Cut Line pada setiap sumbu. Pasangan rectangle bersudut runcing memakai jalur penanganan tersendiri.

Saat beberapa Design memakai satu Cut Line bersama tanpa Quantity `/n`, susunan Design pertama yang aktif menjadi template untuk Design lain pada container terpisah. Satu susunan Cut Line dibuat untuk dipakai bersama. Quantity Design pertama menjadi acuan template; Quantity `0` tetap melewati Design terkait.

### Minimum

`optMinimum` membandingkan rencana **Straight** dan **ZigZag** secara keseluruhan, dengan prioritas:

1. Jumlah area paling sedikit.
2. Jumlah Design paling banyak.
3. Isi area awal lebih banyak.
4. **Straight** jika hasil tetap sama.

Rectangle, termasuk rectangle dengan sudut lengkung, memakai Straight. Jika seluruh acuan aktif berupa rectangle, perbandingan ZigZag dilewati.

Perencanaan memeriksa batas area, zona mark, dan collision. Saat susunan perlu menyesuaikan posisi terhadap mark, Minimum mencoba pergeseran bersama sebelum mengurangi isi susunan.

### Medium

`optMedium` mencoba kandidat rotasi dan pergeseran untuk mencari susunan yang lebih baik. Rencana Minimum yang lolos validasi juga ikut menjadi kandidat.

- Bentuk non-rectangle mencoba rotasi relatif terhadap sumber dengan langkah **45°**: `0°`, `45°`, `90°`, `135°`, `180°`, `225°`, `270°`, dan `315°`.
- Rectangle memakai rotasi dengan langkah **90°**.
- Segitiga dan kurva custom sederhana yang memenuhi syarat dapat mencoba pola pasangan dengan rotasi berlawanan, termasuk offset dan perapatan menurut kontur.
- Pemilihan mempertimbangkan kapasitas, jumlah area, keteraturan pola, dan luas bounding box acuan collision saat hasil utama sama.
- Perapatan dan perataan tetangga tetap diperiksa terhadap gap, collision, batas area, dan mark.

Medium memakai pencarian heuristik dengan batas kandidat. Hasilnya bukan jaminan nesting paling optimal untuk setiap bentuk. Waktu Process dapat bertambah pada kurva kompleks, banyak sumber, atau pencarian Quantity kelompok.

### Model Area

| Model | Kontrol | Area susun | Acuan sheet untuk mark |
| --- | --- | --- | --- |
| KissA | `optKissA` | **320 × 470 mm** | **325 × 485 mm** |
| DieA | `optDieA` | **320 × 470 mm** | **325 × 485 mm** |
| Custom | `optCustom` | `txbAreaWidth` × `txbAreaHeight`; default **320 × 470 mm** | Tanpa zona mark preset |

Kolom ukuran area hanya aktif pada Custom. Lebar dan tinggi harus lebih besar dari nol.

Zona mark preset menjadi area yang dihindari saat menyusun objek:

- **KissA:** empat zona persegi **10 × 10 mm** dan satu zona persegi **12 × 12 mm**.
- **DieA:** empat zona ellipse **11 × 11 mm**, dengan pusat `(11; 11)`, `(11; 474)`, `(314; 11)`, dan `(314; 474)` dalam koordinat sheet preset, satuan mm.

Pemeriksaan zona mark memakai toleransi penetrasi **0,2 mm** dari tepi mark. Zona ini adalah acuan perencanaan; Process tidak membuat objek sensor atau mark fisik.

Ukuran sheet preset dipakai sementara saat Process, lalu **ukuran halaman asli dipulihkan**. Hasil tetap mengacu pada pusat halaman asli.

### Gap Horizontal dan Vertical

Gap diatur melalui `txbGapHorizontal` dan `txbGapVertical`. Kolom kosong memakai default sesuai model dan pola:

| Model | Gap H default | Gap V Straight default | Langkah V ZigZag default |
| --- | --- | --- | --- |
| KissA | **1 mm** | **1 mm** | **88%** tinggi objek acuan |
| DieA | **1,5 mm** | **1,5 mm** | **89%** tinggi objek acuan |
| Custom | **0 mm** | **0 mm** | **88%** tinggi objek acuan |

Pada KissA/DieA, rectangle bersudut runcing memakai gap default **0 mm**. Input manual diutamakan.

Objek acuan adalah Cut Line pada pasangan, atau Design pada standalone.

**Minimum** mendukung `chkHPercent` dan `chkVPercent`. Nilai persen menentukan **langkah penempatan** berdasarkan lebar atau tinggi objek acuan, bukan persentase celah kosong. Misalnya, `100%` berarti langkah sebesar satu ukuran acuan. Nilai persen harus lebih besar dari nol.

**Medium** memakai gap kontur dalam mm dan menerima nilai minimal **0 mm**. Persen dinonaktifkan; sisi miring memakai nilai terbesar antara Gap H dan Gap V. Kolom kosong tetap mengikuti default model, dengan penyesuaian rectangle bersudut runcing.

Input gap Minimum dan Medium disimpan terpisah saat berpindah mode. Angka abu-abu pada kolom kosong adalah placeholder, bukan nilai manual.

### Quantity

`txbQuantity` berlaku untuk Design. Jika beberapa Design dipilih di `lbxObjects`, perubahan Quantity diterapkan ke seluruh Design terpilih.

| Nilai | Perilaku |
| --- | --- |
| Kosong | Mengisi sisa area aktif; mencoba area baru jika sumber tidak muat pada sisa area. |
| `0` | Melewati Design tersebut. |
| Bilangan bulat positif, misalnya `100` | Menyusun jumlah yang diminta; kebutuhan yang belum tertampung dilanjutkan ke area berikutnya. |
| `/n`, misalnya `/1` | Mengelompokkan Design dengan nomor yang sama untuk dibagi dalam satu container. |

Tanda `-` yang tampil saat Quantity kosong adalah placeholder; gunakan kolom kosong untuk pengisian otomatis. Token `/` harus disertai nomor kelompok positif; `/` saja dan `/0` tidak valid.

**Quantity kelompok `/n`** tersedia pada Minimum dan Medium, untuk Design standalone, Cut Line bersama, maupun Cut Line masing-masing. Contohnya, beri `/1` pada dua Design agar keduanya mengisi satu container kelompok yang sama. `/2` membuat kelompok lain.

Kelompok diproses menurut nomor, lalu Design dengan Quantity biasa diproses terpisah. Pembagian memakai template jika acuan sesuai, atau mencari kuota yang muat dalam satu area. Pembagian template memiliki selisih maksimal satu objek; pencarian kuota dan surplus mengizinkan selisih maksimal **dua objek** antar Design. Nomor setelah `/` adalah **ID kelompok**, bukan jumlah salinan atau pembagi.

Susunan Cut Line bersama yang identik dapat memakai satu template Cut Line. Kelompok `/n` tetap harus muat dalam satu container.

### Hasil, Layer, dan Undo

- Area pertama mengacu pada pusat halaman asli. Area tambahan ditempatkan ke kanan pada **halaman yang sama**, dengan jarak antartepi area **50 mm**.
- Objek sumber digunakan sebagai hasil pertama, lalu diduplikasi untuk slot berikutnya.
- Hasil digroup menurut area, sumber, peran, dan layer; Design dan Cut Line tetap terpisah.
- Layer sumber dipertahankan dengan koreksi urutan pasangan bila diperlukan. Pada layer bernama `Layer n`, pasangan yang terbalik dapat ditukar layernya; Cut Line pada layer yang sama ditempatkan di depan Design.
- Z-Order anggota group mengikuti urutan slot dalam rencana.
- Group hasil menjadi selection aktif. Process mengaktifkan **Pick Tool**, termasuk jika dijalankan saat **Shape Tool** aktif.
- Operasi berada dalam satu `CommandGroup`; gunakan **Undo sekali** sebelum mencoba mode lain dengan daftar sumber yang sama.
- Satuan dokumen dan ruler yang dipakai sementara dipulihkan setelah Process.

## Cara menggunakan

1. Buka dokumen dan halaman sumber di CorelDRAW, lalu buka `AutoAlignMenu`.
2. Pilih objek Design dan tekan `cmdSetDesign`.
3. Jika memakai Cut Line, daftarkan satu Cut Line bersama atau pasangan sebanyak Design melalui `cmdSetCutLine`. Periksa urutan pasangan dalam daftar.
4. Pilih KissA, DieA, atau Custom. Untuk Custom, isi ukuran area jika ingin mengganti default.
5. Pilih Minimum atau Medium, lalu atur Gap H/V sesuai kebutuhan.
6. Pilih Design di `lbxObjects` dan atur Quantity. Gunakan Ctrl/Shift untuk menerapkan jumlah atau `/n` ke beberapa Design sekaligus.
7. Tekan `cmdProcess`, lalu periksa susunan, jumlah, gap, dan pasangan hasil.
8. Untuk membandingkan mode lain, lakukan Undo sekali sebelum Process berikutnya. `cmdClose` menutup form.

## Integrasi dan Source

Repository menyediakan source VBA berupa code-behind UserForm, class module, dan standard module. Layout visual UserForm lengkap serta paket GMS siap pakai tidak disertakan. Siapkan UserForm `AutoAlignMenu` dan kontrol sesuai nama dalam kode saat memasangnya di project VBA CorelDRAW.

| Lokasi | Peran |
| --- | --- |
| `src/forms/AutoAlignMenu.vba` | Pendaftaran objek, pilihan model/mode, input gap, Quantity, Process, dan callback MacroRunner. |
| `src/classes/AAPresenter.cls`, `AAObjectStore.cls`, `AASettings.cls` | Penghubung form, daftar objek, dan validasi pengaturan. |
| `src/classes/AAProcessEngine.cls`, `AAResultBuilder.cls` | Menjalankan proses, menangani pasangan/kelompok, dan membangun hasil. |
| `src/classes/AAPlanner.cls`, `AAMediumPlanner.cls`, `AAQuantityPlanner.cls` | Perencanaan Minimum, Medium, dan kuota kelompok. |
| `src/classes/AAModel.cls`, `AAGeometry.cls`, `AAContour.cls`, `AAOccupiedIndex.cls` | Acuan bentuk, kontur, rotasi, gap, dan pemeriksaan tetangga. |
| `src/modules/AAAlgorithm.bas` | Konstanta serta aturan layer dan Z-Order. |
| `src/modules/AADebug.bas` | Debug, profiling, dan diagnostik memori. |
| `src/modules/MRTargetBridge.bas` | Membuka form melalui `OpenMacro` dan menghubungkan callback MacroRunner. |
| `Changelog.log` | Riwayat pengembangan dan catatan pengujian. |

Integrasi MacroRunner yang tersedia menghubungkan pembukaan/penutupan form serta pemberitahuan waktu mulai dan selesai Process. Source ini belum menyediakan `ValidateBehavior` atau `RunBehavior` untuk MacroBehavior otomatis.

### Debug Opsional

Debug nonaktif secara default. Jalankan perintah berikut di **Immediate Window** pada project VBA AutoAlign:

```vb
AADebugOn           ' Ringkas
AADebugOn True      ' Detail
AADebugFileOn       ' Tambahkan log ke TEMP\AutoAlign-Debug.txt
AADebugFileOff      ' Hentikan keluaran file
AADebugOff          ' Matikan debug dan tutup file log
```

Debug menyediakan informasi kandidat, waktu tahap Process, pemeriksaan gap/overlap, penempatan, dan snapshot memori proses CorelDRAW. `AADebugFileOn` juga menerima path file dan menambahkan log ke isi yang sudah ada.

## Status dan Batasan

- Mode yang aktif adalah **Minimum** dan **Medium**. `optMaximum` dinonaktifkan dan `cmdOptimize` belum memiliki implementasi pemrosesan.
- Perencanaan dibatasi hingga **5000 objek per sumber**; satu Cut Line bersama juga mempunyai batas jumlah terkait.
- Sumber harus mempunyai ukuran yang valid dan kurva yang dapat dibaca. Pada group yang perlu dibaca geometrinya, anggota tanpa kurva dapat menyebabkan validasi gagal.
- Daftar mempertahankan ID sumber. Jika sumber dihapus, dipindah, atau berubah struktur, pulihkan sumber atau daftarkan ulang sebelum Process.
- Hasil bergantung pada bentuk, jumlah sumber, model area, gap, Quantity, dan batas pencarian. Catatan pengujian pada Changelog mengacu pada kasus yang disebutkan di sana.

## Lisensi dan Masukan

Project ini menggunakan **MIT License**; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
