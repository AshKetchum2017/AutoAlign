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

Untuk pasangan selain rectangle bersudut runcing, Design boleh lebih kecil maksimal **0,2 mm** atau lebih besar maksimal **2 mm** dari Cut Line pada setiap sumbu.

Saat beberapa Design memakai satu Cut Line bersama tanpa Quantity `/n`, susunan Design pertama yang aktif menjadi template untuk Design lain pada container terpisah. Satu susunan Cut Line dibuat untuk dipakai bersama. Quantity Design pertama menjadi acuan template; Quantity `0` tetap melewati Design terkait.

### Minimum

`optMinimum` membandingkan rencana **Straight** dan **ZigZag** secara keseluruhan, dengan prioritas:

1. Jumlah area paling sedikit.
2. Jumlah Design paling banyak.
3. Isi area awal lebih banyak.
4. **Straight** jika hasil tetap sama.

Rectangle, termasuk rectangle dengan sudut lengkung, disusun dengan pola Straight.

Perencanaan memeriksa batas area, zona mark, gap kontur, dan collision. Langkah Straight dan ZigZag mengikuti kontur acuan serta gap yang diminta. Saat susunan perlu menyesuaikan posisi terhadap mark, Minimum mencoba pergeseran bersama sebelum mengurangi isi susunan.

Untuk Quantity berupa angka dengan beberapa Design berbeda dalam satu container, Minimum juga mencoba posisi tambahan berdasarkan tepi dan alignment tetangga pada **rotasi 0°** sebelum membuka container baru. Pencarian tambahan ini tetap dibatasi dan harus memenuhi gap, batas area, serta zona mark.

### Medium

`optMedium` mencoba kandidat rotasi dan pergeseran untuk mencari susunan yang lebih baik. Rencana Minimum yang lolos validasi juga ikut menjadi kandidat.

- Bentuk non-rectangle mencoba rotasi relatif terhadap sumber dengan langkah **45°**: `0°`, `45°`, `90°`, `135°`, `180°`, `225°`, `270°`, dan `315°`.
- Rectangle memakai rotasi dengan langkah **90°**.
- Segitiga dan kurva custom yang memenuhi syarat dapat memakai pola pasangan dengan rotasi berlawanan, offset vertikal, dan perapatan menurut kontur.
- Pemilihan mengutamakan jumlah area paling sedikit, jumlah Design paling banyak, lalu keteraturan pola. Setelah itu, luas bounding box acuan collision ikut dibandingkan.
- Perapatan dan perataan tetangga tetap diperiksa terhadap gap, collision, batas area, dan mark.

Jika lebar dan tinggi susunan acuan collision pada **setiap area** setara dalam toleransi **0,001 mm**, Medium memilih total rotasi tambahan paling sedikit terhadap orientasi sumber. Toleransi ini hanya untuk membandingkan kandidat, bukan kelonggaran gap atau batas area. Jika total rotasi juga sama, isi area awal diprioritaskan.

Medium memakai pencarian heuristik dengan batas kandidat. Hasilnya bukan jaminan nesting paling optimal untuk setiap bentuk. Waktu Process dapat bertambah pada kurva kompleks, banyak sumber, atau pencarian Quantity kelompok.

### Model Area

| Model | Kontrol | Area susun | Acuan sheet untuk mark |
| --- | --- | --- | --- |
| KissA | `optKissA` | Preferensi; awal **320 × 470 mm** | **325 × 485 mm** |
| DieA | `optDieA` | Preferensi; awal **320 × 470 mm** | **325 × 485 mm** |
| Custom | `optCustom` | `txbAreaWidth` × `txbAreaHeight`; default **320 × 470 mm** | Tanpa zona mark preset |

Kolom ukuran area hanya aktif pada Custom. Lebar dan tinggi harus lebih besar dari nol.

Zona mark preset menjadi area yang dihindari saat menyusun objek:

| Model | Bentuk dan ukuran zona | Pusat zona pada sheet preset, dalam mm |
| --- | --- | --- |
| KissA | Empat persegi **10 × 10 mm** | `(6; 10)`, `(6; 475)`, `(319; 10)`, `(319; 475)` |
| KissA | Satu persegi **12 × 12 mm** | `(162,5; 11,5)` |
| DieA | Empat ellipse **11 × 11 mm** | `(11; 11)`, `(11; 474)`, `(314; 11)`, `(314; 474)` |

Pemeriksaan zona mark memakai toleransi penetrasi **0,2 mm** dari tepi mark. Zona ini adalah acuan perencanaan; Process tidak membuat objek sensor atau mark fisik.

Ukuran sheet preset dipakai sementara saat Process, lalu **ukuran halaman asli dipulihkan**. Hasil tetap mengacu pada pusat halaman asli.

### Gap Horizontal dan Vertical

Gap diatur melalui `txbGapHorizontal` dan `txbGapVertical`. Isi jarak kontur dalam **mm**, dengan nilai minimal **0 mm**. Pengaturan ini berlaku pada Minimum dan Medium.

Kolom kosong memakai preferensi default model. Nilai bawaan:

| Model | Gap H default | Gap V default |
| --- | --- | --- |
| KissA | **0 mm** | **0 mm** |
| DieA | **0 mm** | **0 mm** |
| Custom | **0 mm** | **0 mm** |

Pada KissA/DieA, rectangle bersudut runcing memakai gap default **0 mm**. Input manual diutamakan. Angka desimal dapat menggunakan titik atau koma, misalnya `1.5` atau `1,5`.

Tombol `cmdDefaultSetting` membuka **DefaultSettings** untuk mengatur Area W/H dan Gap H/V per model. Berpindah KissA/DieA mempertahankan draft masing-masing; **Save** menyimpan keduanya, sedangkan **Cancel** atau tombol X membuang perubahan. Area harus positif, gap minimal 0 mm, dan seluruh kolom harus terisi. Hint form utama mengikuti preferensi setelah Save; input gap manual dan input area Custom dipertahankan.

Preferensi disimpan di registry Windows dan dimuat saat AutoAlign digunakan kembali. Ukuran halaman preset dan koordinat mark terpisah dari pengaturan area susun.

Objek acuan adalah Cut Line pada pasangan, atau Design pada standalone. Pemilihan gap mengikuti arah antarpusat bounding box acuan collision:

| Posisi relatif acuan | Gap yang dipakai |
| --- | --- |
| Sebaris | Gap H |
| Sekolom | Gap V |
| Diagonal | Nilai terbesar antara Gap H dan Gap V |

Gap diperiksa terhadap kontur acuan sesuai arah tetangga pada tabel. Gap `0` memperbolehkan objek bersinggungan, dengan pemeriksaan collision tetap aktif. Pada Minimum, langkah Straight dan ZigZag mengikuti kontur serta gap yang diminta.

Input gap Minimum dan Medium disimpan terpisah saat berpindah mode. Pada kunjungan pertama ke Medium, input mm dari Minimum diteruskan sebagai nilai awal. Angka abu-abu pada kolom kosong adalah placeholder, bukan nilai manual.

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

### Indikator Progres Process

Indikator progres tampil tanpa perlu mengaktifkan Debug:

| Kontrol | Tampilan |
| --- | --- |
| `lblProgressStage` | Tahap yang sedang berjalan, seperti membaca geometri, mencari susunan, merapatkan, atau mengelompokkan hasil. |
| `lblProgressDetail` | Detail pekerjaan, kelompok Quantity, nomor percobaan/pola, atau jumlah yang sudah diproses. |
| `lblProgressTime` | Durasi Process dalam format menit:detik. |
| `frmProgress` dan `lblProgressFill` | Bar aktivitas atau progres tahap yang sedang berjalan. |

Saat pencarian susunan, bar bergerak sebagai indikator aktivitas; bar ini **bukan persentase keseluruhan Process**. Pada tahap yang mempunyai total, seperti membaca geometri dan membuat objek hasil, isi bar mengikuti jumlah yang telah diproses.

Selama Process, input form dan interaksi dengan jendela dokumen dinonaktifkan sementara. Tombol **Close** berubah menjadi **Cancel** untuk meminta pembatalan. Permintaan diproses pada checkpoint berikutnya; operasi CorelDRAW yang sedang berjalan perlu selesai terlebih dahulu. Selama pemulihan, tombol menampilkan **Membatalkan...**. Setelah pembatalan, perubahan Process dipulihkan, form tetap terbuka, dan tombol kembali menjadi **Close**. Saat tidak ada proses berjalan, **Close** menutup form.

### Hasil, Layer, dan Undo

- Area pertama mengacu pada pusat halaman asli. Area tambahan ditempatkan ke kanan pada **halaman yang sama**, dengan jarak antartepi area **50 mm**.
- Objek sumber digunakan sebagai hasil pertama, lalu diduplikasi untuk slot berikutnya.
- Hasil digroup menurut area, sumber, peran, dan layer; Design dan Cut Line tetap terpisah.
- Layer sumber dipertahankan dengan koreksi urutan pasangan bila diperlukan. Pada layer bernama `Layer n`, pasangan yang terbalik dapat ditukar layernya; Cut Line pada layer yang sama ditempatkan di depan Design.
- Urutan anggota group mengikuti urutan penempatan.
- Group hasil menjadi selection aktif. Process mengaktifkan **Pick Tool**, termasuk jika dijalankan saat **Shape Tool** aktif.
- Gunakan **Undo sekali** untuk membatalkan hasil Process sebelum mencoba mode lain dengan daftar sumber yang sama.
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
| `src/classes/AADefaultPreferences.cls`, `src/forms/DefaultSettings.vba` | Preferensi Area/Gap per model, draft editor dan penyimpanan registry. |
| `src/classes/AAProcessEngine.cls`, `AAResultBuilder.cls` | Menjalankan proses, menangani pasangan/kelompok, dan membangun hasil. |
| `src/classes/AAPlanner.cls`, `AAMediumPlanner.cls`, `AAQuantityPlanner.cls` | Perencanaan Minimum, Medium, dan kuota kelompok. |
| `src/classes/AAModel.cls`, `AAGeometry.cls`, `AAContour.cls`, `AAOccupiedIndex.cls` | Acuan bentuk, kontur, rotasi, gap, dan pemeriksaan tetangga. |
| `src/modules/AAAlgorithm.bas` | Konstanta serta aturan layer dan Z-Order. |
| `src/modules/AADebug.bas` | Debug, profiling, dan diagnostik memori. |
| `src/modules/AAProgress.bas` | Indikator tahap/detail/durasi, pembaruan bar, serta penguncian sementara jendela dokumen saat Process. |
| `src/modules/MRTargetBridge.bas` | Membuka form melalui `OpenMacro` dan menghubungkan callback MacroRunner. |
| `Changelog.log` | Riwayat perubahan. |

Integrasi MacroRunner mendukung pembukaan/penutupan form dan pencatatan durasi Process. Pengaturan serta penyusunan dijalankan melalui form AutoAlign.

Preferensi pengguna tersimpan pada `HKCU\Software\VB and VBA Program Settings\RinCorelMacros\AutoAlignDefaults`, nilai `ProfilesV1`.

### Debug Opsional

Helper Debug dan progres disembunyikan dari daftar macro workspace. Debug nonaktif secara default. Jalankan perintah berikut di **Immediate Window** pada project VBA AutoAlign:

```vb
AADebugOn           ' Ringkas
AADebugOn True      ' Detail
AADebugFileOn       ' Tambahkan log ke TEMP\AutoAlign-Debug.txt
AADebugFileOff      ' Hentikan keluaran file
AADebugOff          ' Matikan debug dan tutup file log
```

Debug menyediakan informasi kandidat, waktu tahap Process, pemeriksaan gap/overlap, penempatan, dan snapshot memori proses CorelDRAW. `AADebugFileOn` juga menerima path file dan menambahkan log ke isi yang sudah ada.

## Status dan Batasan

- Process menyediakan mode **Minimum** dan **Medium**. `optMaximum` dinonaktifkan. Optimize bersifat eksperimental; periksa hasil pemadatannya sebelum digunakan.
- Perencanaan dibatasi hingga **5000 objek per sumber**; satu Cut Line bersama juga mempunyai batas jumlah terkait.
- Sumber harus mempunyai ukuran yang valid dan kurva yang dapat dibaca. Pada group yang perlu dibaca geometrinya, anggota tanpa kurva dapat menyebabkan validasi gagal.
- Daftar mempertahankan ID sumber. Jika sumber dihapus, dipindah, atau berubah struktur, pulihkan sumber atau daftarkan ulang sebelum Process.
- Hasil bergantung pada bentuk, jumlah sumber, model area, gap, Quantity, dan batas pencarian.

## Lisensi dan Masukan

Project ini menggunakan **MIT License**; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
