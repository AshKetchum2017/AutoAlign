# AutoAlign

AutoAlign adalah macro VBA untuk CorelDRAW yang dirancang untuk menyusun dan menggandakan objek **Design** beserta **Cut Line** secara otomatis di dalam area produksi.

Macro ini membantu mengurangi pekerjaan manual saat menentukan jumlah objek yang muat, mengatur jarak, mencoba pola susunan, dan menjaga hasil agar tetap berada di dalam area serta menghindari zona mark.

AutoAlign menyediakan mode **Minimum** dan **Medium**, model area **KissA**, **DieA**, dan **Custom**, serta pengaturan **Quantity** untuk jumlah tertentu, pengisian sisa area, atau pembagian beberapa Design dalam satu container. Pasangan Cut Line dan batas rotasi diatur per Design; susunan yang sudah ada dapat dioptimasi atau diperiksa melalui operasi tersendiri.

## Pilih Operasi yang Sesuai

| Operasi | Input di `lbxObjects` | Perilaku |
| --- | --- | --- |
| **Process** (`cmdProcess`) | Design sumber dan Cut Line jika digunakan. | Membuat susunan dari sumber, menggandakan objek sesuai Quantity, dan mengelompokkan hasil. Memerlukan sedikitnya satu Design. |
| **Optimize** (`cmdOptimize`) | Group container susunan yang sudah ada. | Menggeser anggota untuk menyesuaikan gap atau merapatkan susunan, dengan jumlah dan rotasi tetap. Mendukung container Cut Line saja. |
| **Periksa Susunan** (`cmdCheckLayout`) | Container hasil atau objek individual yang ingin diperiksa. | Membaca gap, collision, batas area, zona mark, dan pasangan. Tidak memperbaiki atau mengubah geometri. Mendukung Cut Line saja. |

**Daftar Objects menentukan target ketiga operasi.** Selection Corel dipakai saat pendaftaran; berpindah selection setelah itu tidak mengganti isi daftar. Setelah Process, daftar sumber tidak otomatis menjadi daftar seluruh hasil. Untuk Optimize atau Periksa Susunan, kosongkan daftar dan daftarkan container hasil yang ingin ditangani, lalu tetapkan pasangan container-nya.

## Alur Singkat

1. Aktifkan dokumen dan halaman sumber, lalu buka `AutoAlignMenu`.
2. Daftarkan objek melalui `cmdSetDesign` dan, jika digunakan, `cmdSetCutLine`.
3. Pilih Design pada daftar, lalu tentukan **Pasangan Cut Line**, **Quantity**, dan **Rotasi Design** sesuai kebutuhan.
4. Pilih **KissA**, **DieA**, atau **Custom**, kemudian **Minimum** atau **Medium**. Atur Area dan Gap H/V.
5. Jalankan `cmdProcess`, tunggu sampai selesai, lalu periksa susunan pada dokumen.
6. Untuk pemeriksaan hasil, daftarkan container hasil dan jalankan **Periksa Susunan**. Untuk mencoba Process lain dengan daftar sumber yang sama, lakukan **Undo sekali** sebelum memproses kembali.

Nilai awal form adalah **KissA** dan **Minimum**. **Maximum** (`optMaximum`) dinonaktifkan. Optimize masih dalam tahap pengembangan; periksa hasilnya sebelum digunakan untuk produksi.

## Main Features

### Object Container

Objek yang akan diproses didaftarkan ke `lbxObjects` sebagai:

- **Design** melalui `cmdSetDesign`.
- **Cut Line** melalui `cmdSetCutLine`.

Daftar menyimpan peran, `StaticID`, ID pasangan Cut Line, Quantity, dan pilihan rotasi setiap Design. Objek dengan ID yang sama tidak ditambahkan dua kali.

Pendaftaran beberapa objek sekaligus mengikuti urutan baris dari atas ke bawah, lalu kiri ke kanan. Baris ditentukan dari pusat Y tertinggi dengan toleransi **1 mm** terhadap objek pertama pada baris tersebut.

`lbxObjects` mendukung pilihan banyak baris menggunakan **Ctrl/Shift**. `cmdRemove` menghapus baris terpilih dan `cmdClear` mengosongkan daftar; keduanya mengubah daftar tanpa menghapus objek di dokumen.

Semua objek dalam daftar harus berasal dari **satu dokumen dan satu halaman**, berada langsung pada halaman atau berupa group induk, serta berada pada layer yang dapat diedit.

### Design dan Cut Line

AutoAlign mendukung tiga susunan input:

| Input | Perilaku |
| --- | --- |
| Design tanpa Cut Line | Susunan memakai siluet luar Design sebagai acuan. |
| Satu Cut Line | Semua Design dipasangkan ke ID Cut Line yang sama. |
| Cut Line sebanyak Design | Setiap Design dipasangkan ke satu ID Cut Line berbeda yang dipilih melalui Pasangan Cut Line. |

Setelah mendaftarkan objek, pilih baris Design lalu tentukan **Pasangan Cut Line** melalui `cmbCutPartner`. Hubungan disimpan berdasarkan ID dalam dokumen sumber. Contoh ringkasan daftar: `Design | Obj. ID: /19 -> Cut /26`. Pendaftaran tidak memilih pasangan secara otomatis; urutan baris tidak menentukan pasangan.

- **Ctrl/Shift** dapat memilih beberapa Design untuk diberi pasangan yang sama; baris Cut Line diabaikan. Pilihan berbeda ditampilkan sebagai **Campuran** tanpa mengubah hubungan sampai pengguna memilih Cut Line atau **Tanpa Cut Line**.
- Untuk susunan tanpa pasangan, pilih **Tanpa Cut Line** pada semua Design dan hapus semua baris Cut Line dari daftar. Campuran Design standalone dan berpasangan, beberapa kelompok yang berbagi Cut Line berbeda, serta Cut Line tidak terpakai belum didukung.
- Menghapus baris Cut Line mempertahankan ID pasangan pada Design dan menampilkan **Pasangan hilang /ID**. Pilih ulang pasangan atau lepaskan secara eksplisit. Remove/Clear hanya mengubah daftar, bukan objek dokumen. Quantity dan rotasi tetap mengikuti Design.
- **Pilih Pasangan** (`cmdSelectPair`) aktif ketika tepat satu Design dipilih. Tombol menyeleksi Design beserta Cut Line-nya pada dokumen/halaman sumber, atau Design saja bila tanpa pasangan. Jika salah satu ID hilang, seleksi dokumen dipertahankan.
- Process, Optimize, dan Periksa Susunan memvalidasi hubungan yang sama. Pasangan belum dipilih, hilang, atau di luar tiga susunan ditolak sebelum pemrosesan. Untuk Optimize/Periksa Susunan, daftarkan container hasil dan pilih ulang hubungan container tersebut. Keduanya juga mendukung daftar yang hanya berisi Cut Line.
- Hubungan berlaku selama daftar aktif; Clear, penghapusan Design, atau penutupan form membuang hubungan terkait. Hubungan tidak disimpan ke registry atau dokumen.

Untuk pasangan, **Cut Line menjadi acuan pola, gap, dan collision antarunit**. Bounds Design tetap diperiksa agar muat di dalam area. Artwork Design berpasangan dapat bertumpuk selama acuan Cut Line memenuhi aturan susunan.

Untuk pasangan selain rectangle bersudut runcing, Design boleh lebih kecil maksimal **0,2 mm** atau lebih besar maksimal **2 mm** dari Cut Line pada setiap sumbu.

Saat beberapa Design memakai satu Cut Line bersama tanpa Quantity `/n`, susunan Design pertama yang aktif menjadi template untuk Design lain pada container terpisah. Satu susunan Cut Line dibuat untuk dipakai bersama. Quantity Design pertama menjadi acuan template; Quantity `0` tetap melewati Design terkait.

### Batas rotasi per Design

Pilih Design pada `lbxObjects`, lalu atur `cmbRotation`. Ctrl/Shift dapat dipakai untuk menerapkan pilihan yang sama pada beberapa Design. Baris Cut Line diabaikan dan mengikuti rotasi Design pasangannya. Ringkasan `Rot:` pada daftar menunjukkan pilihan setiap Design.

| Pilihan | Sudut tambahan yang diizinkan pada Process Medium |
| --- | --- |
| Otomatis | Mengikuti kandidat AA: langkah 45° untuk non-rectangle dan 90° untuk rectangle. |
| Tanpa rotasi | 0°. |
| 0 / 180 | 0° dan 180°. |
| 90 / 270 | 90° dan 270°. |
| Kelipatan 90 | 0°, 90°, 180°, dan 270°. |

Sudut relatif terhadap **orientasi sumber saat Process**, bukan sudut absolut terhadap halaman. Rotasi sumber 30° dengan pilihan 0 / 180 berarti orientasi hasil 30° atau 210°. Pembatasan berlaku pada pola utama, pola pasangan berlawanan, pencarian kuota, serta pengisian sisa ruang; jumlah hasil dapat berkurang karena pilihan orientasi lebih sedikit.

- Default saat pendaftaran adalah **Otomatis**. Pilihan mengikuti baris Design selama daftar aktif; Remove/Clear atau penutupan form membuang pilihan terkait. Pengaturan ini tidak disimpan ke registry.
- Jika beberapa Design terpilih memiliki pilihan berbeda, ComboBox kosong dan label menampilkan **Rotasi Design (Campuran)**. Tidak ada perubahan sampai pengguna memilih preset. Tanpa Design terpilih, kontrol dinonaktifkan.
- Untuk **satu Cut Line bersama**, semua Design aktif memakai irisan sudut yang diizinkan. Contoh: Otomatis + 0 / 180 memakai 0°/180°; Tanpa rotasi + Kelipatan 90 memakai 0°. Design dengan Quantity 0 tidak membatasi irisan. Aturan ini berlaku juga bila pemakai Cut Line bersama berada pada kelompok `/n` berbeda. Pilihan asli setiap Design di daftar tetap tersimpan. Pada Medium, jika irisan kosong, misalnya 0 / 180 bersama 90 / 270, Process menolak kombinasi tersebut dan meminta pilihan rotasi yang kompatibel.
- **Minimum** tetap memakai 0° dan mempertahankan jalur template Minimum. **Optimize** mempertahankan rotasi susunan yang sudah ada. **Periksa Susunan** tidak mengaudit preset rotasi terhadap riwayat sumber.

### Minimum

`optMinimum` membandingkan rencana **Straight** dan **ZigZag** secara keseluruhan, dengan prioritas:

1. Jumlah area paling sedikit.
2. Jumlah Design paling banyak.
3. Isi area awal lebih banyak.
4. **Straight** jika hasil tetap sama.

Rectangle, termasuk rectangle dengan sudut lengkung, disusun dengan pola Straight.

Perencanaan memeriksa batas area, zona mark, gap kontur, dan collision, tanpa rotasi tambahan terhadap orientasi sumber. Langkah Straight dan ZigZag mengikuti kontur acuan serta gap yang diminta. Saat susunan perlu menyesuaikan posisi terhadap mark, Minimum mencoba pergeseran bersama sebelum mengurangi isi susunan.

Untuk Quantity berupa angka dengan beberapa Design berbeda dalam satu container, Minimum juga mencoba posisi tambahan berdasarkan tepi dan alignment tetangga pada **rotasi 0°** sebelum membuka container baru. Pencarian tambahan ini tetap dibatasi dan harus memenuhi gap, batas area, serta zona mark.

### Medium

`optMedium` mencoba kandidat rotasi dan pergeseran untuk mencari susunan yang lebih baik. Rencana Minimum yang lolos validasi juga ikut menjadi kandidat.

- Pada pilihan rotasi Otomatis, bentuk non-rectangle mencoba rotasi relatif terhadap sumber dengan langkah **45°**: `0°`, `45°`, `90°`, `135°`, `180°`, `225°`, `270°`, dan `315°`.
- Rectangle memakai rotasi dengan langkah **90°**.
- Segitiga dan kurva custom yang memenuhi syarat dapat memakai pola pasangan dengan rotasi berlawanan, offset vertikal, dan perapatan menurut kontur. Kurva kompleks dengan estimasi kapasitas grid kecil juga dapat mencoba keluarga 0°/180° dan 90°/270° dalam jumlah kandidat terbatas; pemeriksaan gap memakai kontur detail.
- Pemilihan mengutamakan jumlah area paling sedikit, jumlah Design paling banyak, lalu keteraturan pola. Setelah itu, luas bounding box acuan collision ikut dibandingkan.
- Perapatan dan perataan tetangga tetap diperiksa terhadap gap, collision, batas area, dan mark.

Jika lebar dan tinggi susunan acuan collision pada **setiap area** setara dalam toleransi **0,001 mm**, Medium memilih total rotasi tambahan paling sedikit terhadap orientasi sumber. Toleransi ini hanya untuk membandingkan kandidat, bukan kelonggaran gap atau batas area. Jika total rotasi juga sama, isi area awal diprioritaskan.

Pada satu Design rectangle dengan Quantity kosong dan satu area, Medium juga mencoba mengisi celah yang aman setelah pemusatan. Penambahan ini mempertahankan posisi objek yang sudah tersusun dan tidak memperbesar susunan.

Medium memakai pencarian heuristik dengan batas kandidat. Kandidat pasangan satu area yang jumlah objeknya lebih sedikit dari rencana valid terbaik satu area dilewati sebelum finalisasi. Kandidat dengan jumlah sama tetap dibandingkan ukuran dan rotasinya. Hasilnya bukan jaminan nesting paling optimal untuk setiap bentuk. Waktu Process dapat bertambah pada kurva kompleks, banyak sumber, atau pencarian Quantity kelompok.

### Model Area

| Model | Kontrol | Area susun | Acuan sheet untuk mark |
| --- | --- | --- | --- |
| KissA | `optKissA` | Preferensi; awal **320 × 470 mm** | **325 × 485 mm** |
| DieA | `optDieA` | Preferensi; awal **320 × 470 mm** | **325 × 485 mm** |
| Custom | `optCustom` | `txbAreaWidth` × `txbAreaHeight`; default **320 × 470 mm** | Tanpa zona mark preset |

Kolom ukuran area hanya aktif pada Custom. Lebar dan tinggi harus lebih besar dari nol.

### Margin dan area terlarang Custom

Pilih model **Custom**, isi W/H pada form utama, lalu tekan `cmdCustomMarginSettings` untuk membuka **CustomMarginSettings**. Form utama disembunyikan sementara; editor modeless memungkinkan pengguna memilih objek langsung di CorelDRAW. **Terapkan** menerima seluruh draft, sedangkan **Batal** atau X membuang perubahan dan kembali ke form utama.

- **Margin kiri, kanan, atas, bawah** dalam mm membatasi bounds Design dan Cut Line. W/H tetap ukuran area keseluruhan; jumlah margin horizontal/vertikal harus lebih kecil daripada lebar/tinggi. Semua margin awal bernilai 0.
- **Tambah Rectangle** menerima X, Y, W, H dalam mm. X/Y adalah sudut kiri bawah zona, relatif terhadap kiri bawah area pertama sebelum margin. Area pertama berpusat pada halaman aktif. Lebar/tinggi zona harus positif dan seluruh zona harus berada di dalam area keseluruhan.
- **Ambil Seleksi** mengambil setiap objek terpilih sebagai salinan geometri zona, dengan posisi relatif terhadap area pertama dan orientasi saat diambil. Pilih objek pada dokumen/halaman sumber, lalu tekan tombol tersebut. Kurva harus tertutup dan dapat dibaca oleh AA; satu anggota yang tidak valid membatalkan seluruh penambahan seleksi tersebut. Objek sumber tidak dipindahkan, dihapus, atau diubah menjadi kurva. Perubahan/penghapusan objek sumber setelah pengambilan tidak mengubah zona; hapus zona lalu ambil ulang untuk memperbaruinya.
- Rectangle dan zona dari seleksi masuk ke satu daftar. **Hapus Zona** menghapus baris terpilih; **Kosongkan Zona** menghapus semua zona pada draft tanpa mengubah margin atau objek dokumen. Zona berulang pada setiap area hasil, dengan jarak antararea tetap 50 mm berdasarkan ukuran area penuh.
- CheckBox **Design juga menghindari zona** (`chkIncludeDesign`) awalnya dicentang. Saat aktif, siluet Design dan Cut Line diperiksa terhadap zona. Saat dilepas, pasangan memakai Cut Line saja; Design standalone tetap diperiksa. Checkbox tidak mengubah acuan gap antaritem atau batas margin. Zona memakai batas geometri yang diambil, tanpa penyusutan toleransi mark preset 0,2 mm dan tanpa tambahan Gap H/V; sediakan ruang aman dalam ukuran zona bila diperlukan.
- **Process Minimum/Medium, Optimize, dan Periksa Susunan** memakai pengaturan yang sama. Optimize dengan pembatas Custom aktif memakai grid Process tetap agar margin/zona tidak bergeser mengikuti pusat bbox susunan. Pemeriksaan melaporkan **Area terlarang** untuk zona Custom dan **Batas area** untuk pelanggaran ruang efektif di dalam margin.

Pengaturan terikat dokumen/halaman selama sesi AutoAlign, tetap tersedia saat daftar Objects dikosongkan untuk mendaftarkan container hasil, dan tidak disimpan ke registry atau CDR. Pada dokumen/halaman lain, buka editor lalu Terapkan profil untuk sumber tersebut. Mengubah W/H tidak menskalakan zona; pengaturan diperiksa ulang dan zona di luar area harus dihapus/dibuat ulang. KissA/DieA menggunakan aturan presetnya sendiri.

Pembatas dapat mengurangi jumlah hasil. Pencarian tetap heuristik; banyak zona atau siluet Design kompleks dapat menambah waktu proses. Jika satu template Cut Line bersama tidak aman untuk seluruh Design, proses menolak hasil sebelum pembuatan objek.

Zona mark preset menjadi area yang dihindari saat menyusun objek:

| Model | Bentuk dan ukuran zona | Pusat zona pada sheet preset, dalam mm |
| --- | --- | --- |
| KissA | Empat persegi **10 × 10 mm** | `(6; 10)`, `(6; 475)`, `(319; 10)`, `(319; 475)` |
| KissA | Satu persegi **12 × 12 mm** | `(162,5; 11,5)` |
| DieA | Empat ellipse **11 × 11 mm** | `(11; 11)`, `(11; 474)`, `(314; 11)`, `(314; 474)` |

Koordinat pusat zona diukur dari **kiri bawah sheet preset**. Pada **Process KissA**, pergeseran seluruh kelompok mengutamakan zona mark penuh, termasuk zona tengah 12 × 12 mm. Pencarian mencoba posisi di dalam area nominal dahulu. Jika belum menemukan posisi yang aman, kelompok boleh digeser hingga **0,6 mm melewati setiap sisi area** untuk menghindari mark penuh. Lebar dan tinggi kelompok sendiri harus tetap muat dalam ukuran area nominal; kelonggaran ini tidak dipakai untuk mencari slot atau menambah kapasitas area.

Jika pencarian mark penuh belum menemukan posisi, Process KissA memakai toleransi penetrasi mark **0,2 mm** dengan batas area nominal. DieA tetap memakai toleransi mark 0,2 mm tanpa kelonggaran area. Gap antarobjek, ukuran sheet, dan koordinat mark tidak berubah oleh toleransi ini. Zona ini adalah acuan perencanaan; Process tidak membuat objek sensor atau mark fisik.

**Periksa Susunan** mengenali pergeseran KissA di luar area bila kelompok muat secara nominal, keluar maksimal 0,6 mm per sisi, mark penuh aman, dan mengembalikan kelompok tepat ke dalam area mengenai mark. Pemeriksaan memakai geometri saat ini, bukan riwayat Process. **Optimize** memakai batas area nominal.

Ukuran sheet preset dipakai sementara saat Process, lalu **ukuran halaman asli dipulihkan**. Hasil tetap mengacu pada pusat halaman asli.

### Gap Horizontal dan Vertical

Gap diatur melalui `txbGapHorizontal` dan `txbGapVertical`. Isi jarak kontur dalam **mm**, dengan nilai minimal **0 mm**. Pengaturan ini berlaku pada Minimum dan Medium.

Gunakan angka desimal biasa tanpa satuan atau pemisah ribuan. Titik dan koma diterima sebagai pemisah desimal; notasi eksponen tidak digunakan pada input form.

Kolom kosong memakai preferensi default model. Nilai bawaan:

| Model | Gap H default | Gap V default |
| --- | --- | --- |
| KissA | **0 mm** | **0 mm** |
| DieA | **0 mm** | **0 mm** |
| Custom | **0 mm** | **0 mm** |

Pada KissA/DieA, rectangle bersudut runcing memakai gap default **0 mm**. Input manual diutamakan. Angka desimal dapat menggunakan titik atau koma, misalnya `1.5` atau `1,5`.

Tombol `cmdDefaultSetting` membuka **DefaultSettings** untuk mengatur Area W/H dan Gap H/V pada **KissA dan DieA**. Custom memakai input manual pada form utama dengan nilai bawaan pada tabel. Berpindah KissA/DieA mempertahankan draft masing-masing; **Save** menyimpan keduanya, sedangkan **Cancel** atau tombol X membuang perubahan. Area harus positif, gap minimal 0 mm, dan seluruh kolom harus terisi. Hint form utama mengikuti preferensi setelah Save; input gap manual dan input area Custom dipertahankan.

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

Contoh: `100` meminta 100 salinan Design; `/1` pada dua Design meminta pembagian satu container kelompok, bukan satu salinan per Design.

Tanda `-` yang tampil saat Quantity kosong adalah placeholder; gunakan kolom kosong untuk pengisian otomatis. Token `/` harus disertai nomor kelompok positif; `/` saja dan `/0` tidak valid.

**Quantity kelompok `/n`** tersedia pada Minimum dan Medium, untuk Design standalone, Cut Line bersama, maupun Cut Line masing-masing. Contohnya, beri `/1` pada dua Design agar keduanya mengisi satu container kelompok yang sama. `/2` membuat kelompok lain.

Kelompok diproses menurut nomor, lalu Design dengan Quantity biasa diproses terpisah. Pembagian memakai template jika acuan sesuai, atau mencari kuota yang muat dalam satu area. Pembagian template memiliki selisih maksimal satu objek; pencarian kuota dan surplus mengizinkan selisih maksimal **dua objek** antar Design. Nomor setelah `/` adalah **ID kelompok**, bukan jumlah salinan atau pembagi.

Susunan Cut Line bersama yang identik dapat memakai satu template Cut Line. Kelompok `/n` tetap harus muat dalam satu container.

### Progres dan Cancel

Indikator progres tampil tanpa perlu mengaktifkan Debug:

| Kontrol | Tampilan |
| --- | --- |
| `lblProgressStage` | Tahap yang sedang berjalan, seperti membaca geometri, mencari susunan, merapatkan, atau mengelompokkan hasil. |
| `lblProgressDetail` | Detail pekerjaan, kelompok Quantity, nomor percobaan/pola, atau jumlah yang sudah diproses. |
| `lblProgressTime` | Durasi Process dalam format menit:detik. |
| `frmProgress` dan `lblProgressFill` | Bar aktivitas atau progres tahap yang sedang berjalan. |

Saat pencarian susunan, bar bergerak sebagai indikator aktivitas; bar ini **bukan persentase keseluruhan Process**. Pada tahap yang mempunyai total, seperti membaca geometri dan membuat objek hasil, isi bar mengikuti jumlah yang telah diproses.

Pembaruan panel dan pemrosesan event memakai interval 200 ms pada checkpoint. Operasi CorelDRAW yang sedang berjalan dapat membuat jeda aktual lebih panjang.

Process, Optimize, dan Periksa Susunan memakai panel progres serta mekanisme Cancel yang sama. Selama operasi, input form dan interaksi dengan jendela dokumen dinonaktifkan sementara. Tombol **Close** berubah menjadi **Cancel** untuk meminta pembatalan. Permintaan diproses pada checkpoint berikutnya; operasi CorelDRAW yang sedang berjalan perlu selesai terlebih dahulu. Selama pemulihan, tombol menampilkan **Membatalkan...**. Saat pembatalan, Process mencoba membatalkan perubahan melalui Undo dan Optimize mencoba memulihkan posisi. Form tetap terbuka dan tombol kembali menjadi **Close** setelah cleanup. Jika pemulihan tidak lengkap, ikuti pesan yang meminta pemeriksaan dokumen atau Undo. Periksa Susunan menghasilkan laporan berstatus belum selesai. Saat tidak ada proses berjalan, **Close** menutup form.

### Hasil, Layer, dan Undo

- Area pertama mengacu pada pusat halaman asli. Area tambahan ditempatkan ke kanan pada **halaman yang sama**, dengan jarak antartepi area **50 mm**.
- Objek sumber digunakan sebagai hasil pertama, lalu diduplikasi untuk slot berikutnya.
- Hasil digroup menurut area, sumber, peran, dan layer; Design dan Cut Line tetap terpisah.
- Layer sumber dipertahankan dengan koreksi urutan pasangan bila diperlukan. Jika keduanya bernama `Layer n` dan nomor layer Cut Line lebih kecil, Design/Cut Line ditukar layernya. Pada penamaan lain, layer Cut Line dapat dipindahkan ke atas layer Design. Cut Line pada layer yang sama ditempatkan di depan Design.
- Urutan anggota group mengikuti urutan penempatan.
- Group hasil menjadi selection aktif. Process mengaktifkan **Pick Tool**, termasuk jika dijalankan saat **Shape Tool** aktif.
- Gunakan **Undo sekali** untuk membatalkan hasil Process sebelum mencoba mode lain dengan daftar sumber yang sama.
- Satuan dokumen dan ruler yang dipakai sementara dipulihkan setelah Process.

### Optimize susunan yang sudah ada

`cmdOptimize` membaca anggota group container yang terdaftar di `lbxObjects`. Pilihan `optMinimum` atau `optMedium` menentukan algoritma Optimize. Jumlah, rotasi, struktur group, dan offset pasangan Design/Cut Line dipertahankan; Quantity tidak mengubah susunan pada operasi ini. Container yang hanya berisi Cut Line juga didukung.

**Optimize Minimum** ditujukan untuk susunan baris/kolom sederhana:

1. Catat keanggotaan baris dan kolom dari posisi awal. Pusat bounding box Cut Line menjadi acuan untuk pasangan; tanpa pasangan, gunakan pusat Design. Toleransi baris adalah **25% tinggi anchor**, sedangkan toleransi kolom **25% lebar anchor**. Anggota dibandingkan langsung dengan anchor tetap, tanpa pengelompokan berantai melalui tetangga.
2. Proses baris dari pusat Y terbesar ke terkecil. Anchor tiap baris adalah objek paling kiri; pusat Y anggota mengikuti anchor. Objek berikutnya digeser ke kanan jika gap terlalu kecil atau ke kiri jika terlalu lebar, dengan urutan kiri ke kanan tetap dipakai.
3. Pakai keanggotaan kolom yang sudah dicatat. Anchor paling atas pada tiap kolom mengikuti posisi hasil tahap baris. Pusat X anggota di bawahnya mengikuti anchor; pergeseran ke atas/bawah menyesuaikan gap vertikal. Koreksi ini dapat mengubah pusat Y anggota setelah tahap baris.

Pencarian memakai pemeriksaan gap kontur dan collision yang sama dengan planner, serta memeriksa bounds Design dan zona mark. Gap awal yang terlalu kecil dapat diperbaiki; susunan boleh melebar atau meninggi sampai batas area. Minimum mempertahankan anchor tanpa memusatkan ulang seluruh hasil. Jika pencarian tidak menemukan posisi aman atau mencapai batas pencarian, proses dibatalkan dengan pesan dan posisi objek asli dipertahankan.

**Optimize Medium** merapatkan kelompok baris, kolom, dan objek individual, kemudian mencoba memusatkan hasil jika valid. Susunan awal harus memenuhi area, gap, dan mark. Mode ini dapat dicoba untuk susunan dengan bentuk/ukuran campuran; hasilnya tetap perlu diperiksa. `optMaximum` dinonaktifkan.

Optimize membaca struktur `container → subgroup sumber → instance`; group pada tingkat instance tetap utuh. Pasangan anggota dikenali dari pusat bounds relatif terhadap container dengan toleransi **1 mm**. Untuk Cut Line bersama, container Design harus mempunyai template posisi yang identik dalam toleransi tersebut serta orientasi instance yang sesuai.

**Acuan area Optimize berpusat pada bounding box masing-masing container saat dibaca.** Periksa Susunan memakai grid halaman Process. Karena itu, lolos Optimize tidak berarti susunan pasti lolos pemeriksaan terhadap grid halaman yang dipakai Process.

Satu Optimize memakai satu CommandGroup, lalu memulihkan unit dokumen/ruler serta selection. Seluruh rencana divalidasi sebelum perubahan diterapkan; hasil aktual diperiksa lagi. Jika gagal setelah objek digeser, aplikasi mencoba memulihkan posisi semula. Jika pemulihan tidak lengkap, pesan meminta **Undo sekali**. Cancel tersedia pada checkpoint operasi.

Optimize bersifat eksperimental. Pencarian posisi yang aman dapat berhenti karena batas pencarian; mode ini tidak menjamin susunan paling rapat untuk setiap bentuk.

### Pemeriksaan Susunan

`cmdCheckLayout` (**Periksa Susunan**) membaca susunan yang terdaftar di **Objects AA**, memakai model, Area, dan Gap dari form utama. Quantity serta pilihan Minimum/Medium tidak memengaruhi pemeriksaan. Daftarkan **container hasil susunan** yang ingin diperiksa; daftar sumber sebelum Process tidak otomatis menjadi daftar seluruh hasil.

Pemeriksaan menampilkan temuan **Collision**, **Gap**, **Batas area**, **Zona mark**, serta pasangan yang tidak dapat dikenali. Cut Line menjadi acuan gap/collision ketika tersedia; Design berpasangan diperiksa bounds-nya. Tanpa Cut Line, acuan memakai siluet luar Design. Container yang hanya berisi Cut Line juga dapat diperiksa. Pengecualian default gap rectangle bersudut runcing pada preset tetap mengikuti aturan AA; input gap manual diutamakan.

- Acuan area mengikuti grid Process: area pertama berpusat pada halaman sumber, berikutnya ke kanan dengan jarak antararea **50 mm**. Kolom dipilih dari pusat bounding box container terdekat; origin area tidak dipusatkan ulang mengikuti susunan. Pemeriksa tidak menyimpan nomor area historis: container yang dipindahkan satu kolom penuh akan dianggap berada pada kolom tersebut. Gunakan ukuran area dan posisi halaman yang sama dengan saat Process.
- Struktur yang dibaca adalah `container -> subgroup sumber -> instance`; group pada tingkat instance tetap satu unit. Objek individual non-group juga dapat didaftarkan. Group manual atau struktur yang telah diubah perlu disesuaikan dahulu; pemeriksa tidak menebak kedalaman group sembarang.
- Satu container Cut Line dapat menjadi template bersama. Selain itu, setiap container Design harus dipasangkan ke satu container Cut Line berbeda melalui Pasangan Cut Line. Anggota di dalam pasangan container dicocokkan secara unik melalui pusat bounds relatif terhadap area, dengan toleransi **1 mm**. Ini pencocokan posisi, bukan pembuktian kesamaan bentuk/rotasi atau toleransi gap.
- Pemeriksa hanya membaca objek terdaftar. Gap/collision antarkontur acuan diperiksa di dalam area yang sama, termasuk antarcontainer; objek lain di halaman tidak ikut diperiksa. Template Cut Line bersama diperiksa sekali pada posisi fisiknya.
- Gap memakai validator kontur AA yang sama dengan penyusunan. Temuan menunjukkan kegagalan aturan H/V, bukan angka jarak aktual atau verifikasi independen terhadap algoritma AA.

Hasil tampil pada dialog **Hasil Pemeriksaan Susunan**. Pilih temuan lalu tekan **Pilih Objek** untuk menyeleksi objek terkait; tutup dialog sebelum mengedit, kemudian periksa ulang. Pemeriksaan tidak memindahkan, memutar, mewarnai, atau memperbaiki objek dan tidak membuat CommandGroup. Seleksi awal dan unit dipulihkan sebelum hasil tampil; tombol Pilih Objek sengaja mengganti seleksi.

Ringkasan status, jumlah objek terbaca, pasangan Design/Cut Line yang dikenali, serta pasangan kontur yang diperiksa tampil juga di kotak detail. Jumlah pasangan kontur adalah jumlah pemeriksaan kandidat, bukan seluruh kombinasi objek; filter tetangga dapat melewati pasangan yang jauh.

Cancel, geometri gagal dibaca, ID hilang, daftar tumpang tindih, pasangan ambigu, atau tidak ada objek terbaca menghasilkan status **Pemeriksaan belum selesai**. Daftar dibatasi **1000 temuan pertama**, tetapi hitungan dan pemeriksaan tetap dilanjutkan. Jika debug aktif, label `Inspect`/`Inspect root` mencatat jumlah container, tipe/ID dan jumlah anggotanya, temuan, serta ringkasan hasil.

## Menyiapkan Hasil untuk Optimize atau Pemeriksaan

1. Selesaikan Process dan pertahankan struktur group hasilnya.
2. Tekan `cmdClear` untuk mengosongkan daftar sumber. Tindakan ini tidak menghapus objek di dokumen.
3. Daftarkan container Design hasil melalui `cmdSetDesign`, lalu container Cut Line melalui `cmdSetCutLine` jika digunakan. Untuk Cut Line saja, cukup daftarkan container Cut Line.
4. Pilih setiap Design pada daftar dan tetapkan pasangan **container Cut Line** yang sesuai. Hubungan sumber tidak diteruskan otomatis ke ID container hasil.
5. Pilih model dan Area/Gap yang ingin digunakan. Gunakan **Minimum/Medium + Optimize** untuk menggeser susunan, atau **Periksa Susunan** untuk membaca temuan.
6. Pada dialog hasil pemeriksaan, pilih temuan dan tekan **Pilih Objek**. Tutup dialog sebelum mengedit, lalu jalankan pemeriksaan ulang.

Jangan mendaftarkan container bersamaan dengan anggotanya. Untuk memeriksa banyak area atau seluruh hasil, daftarkan semua container yang diperlukan. Objek yang tidak didaftarkan tidak termasuk cakupan laporan.

## Ringkasan Hasil Produksi

Sesudah **Process** berhasil, tekan **Ringkasan Hasil** (`cmdProductionSummary`) untuk membuka `ProductionSummary`. Dialog tidak muncul otomatis. Ringkasan berisi jumlah Design per sumber dan per area, total Design, total Cut Line fisik, jumlah area hasil, serta total per sumber termasuk Design dengan hasil nol.

Kolom Design/Cut Line memakai **ID sumber saat Process**, bukan ID setiap duplikat. Pada satu Cut Line bersama, hubungan pasangan tetap ditampilkan pada setiap Design, tetapi total Cut Line hanya menghitung objek fisik yang dibuat. Jumlah area menghitung area hasil yang terisi, bukan jumlah lembar bahan. Persentase pemakaian bahan dan limbah tidak dihitung.

Detail mencatat nama dokumen, halaman, waktu selesai, durasi Process hingga pemulihan dokumen, model/mode, ukuran area penuh dan setelah margin, Gap H/V input/default, serta margin/zona Custom. Pengecualian gap default rectangle tajam preset tetap berlaku. Durasi beresolusi detik dan tidak mencakup waktu membaca dialog.

**Salin Ringkasan** menyalin ringkasan, tabel dan detail ke clipboard sebagai teks. Tutup dialog untuk melanjutkan penggunaan AA.

Laporan adalah catatan Process terakhir yang berhasil **selama sesi form AA masih terbuka**. Clear Objects, perubahan input, berpindah dokumen, Optimize, Undo, dan edit manual tidak memperbarui catatan tersebut; periksa nama dokumen dan waktu pencatatannya. Process gagal, dibatalkan, atau tanpa hasil tidak mengganti catatan sukses sebelumnya. Menutup form utama mengakhiri sesi; ringkasan tidak disimpan ke registry atau CDR. Ringkasan bukan hasil Check Layout dan tidak menyatakan susunan sudah bebas collision.

## Preset Pekerjaan Bernama

Tombol **Preset Pekerjaan** (`cmdJobPresets`) membuka `JobPresets`. Preset menyimpan model KissA/DieA/Custom, mode Minimum/Medium, ukuran area penuh, Gap H/V beserta status manual/default, nilai default saat disimpan, dan empat margin Custom. Daftar Objects, pasangan Cut Line, Quantity, rotasi per Design, hasil susunan dan laporan produksi tidak disimpan atau diubah oleh preset.

1. Atur pekerjaan pada form AA; untuk Custom, buka dokumen dan terapkan margin melalui CustomMarginSettings.
2. Buka **Preset Pekerjaan**, isi `txbPresetName`, lalu tekan **Simpan Pengaturan Saat Ini**. Sumber penyimpanan selalu pengaturan form AA, bukan isi preset yang dipilih di daftar.
3. Pilih nama dalam daftar untuk melihat rinciannya. **Terapkan** mengisi pengaturan AA dan menutup dialog; Process tetap dijalankan secara terpisah.
4. **Hapus** menghapus preset terpilih setelah konfirmasi. **Tutup** atau X menutup dialog tanpa menerapkan pilihan. Save/Hapus yang sudah dilakukan tetap tersimpan.

Nama berisi 1–80 karakter setelah spasi tepi dihapus; slash dan karakter kontrol tidak diterima. Nama yang berbeda hanya dalam huruf besar/kecil dianggap sama. Simpan dengan nama yang sama meminta konfirmasi penggantian; gunakan nama baru untuk membuat salinan pengaturan. Preset yang datanya rusak tetap terlihat dan dapat dihapus atau ditimpa, tetapi tidak dapat diterapkan.

**Default yang disimpan tetap berlaku pada model preset tersebut selama sesi AA**, termasuk setelah nilai global DefaultSettings berubah. Status gap default tetap dipertahankan, sehingga aturan default gap 0 untuk rectangle tajam KissA/DieA tidak berubah menjadi gap manual. Process, Optimize, Periksa Susunan, tampilan area dan hint gap memakai nilai yang sama. Input manual tetap diutamakan; gap mode lain yang sudah ada tidak ikut diganti. Saat berpindah ke model lain, default global model itu berlaku; kembali ke model preset memakai default tersimpannya lagi.

Pilih **(Tanpa preset - DefaultSettings)** lalu **Terapkan** untuk melepaskan default pekerjaan yang tersimpan. Model/mode, input manual, margin, zona dan Objects tetap; kolom yang mengikuti default kembali membaca DefaultSettings. Menghapus atau menimpa preset di daftar tidak mengubah salinan yang sudah diterapkan. Preset tidak diterapkan otomatis ketika form AA dibuka kembali.

**Zona terlarang tidak disimpan ke preset.** Save memberi informasi jumlah zona yang tidak disertakan bila pekerjaan Custom aktif memiliki zona. Saat menerapkan preset Custom, zona dan checkbox Design milik dokumen/halaman aktif dipertahankan; empat margin diganti dari preset. Jika ukuran area baru membuat zona keluar batas atau margin tidak menyisakan ruang, penerapan ditolak sebelum pengaturan aktif diganti. Zona dari dokumen/halaman lain tidak dipindahkan ke pekerjaan baru. KissA/DieA mempertahankan profil Custom untuk digunakan kembali ketika model Custom dipilih.

Preset tersimpan per pengguna Windows di `HKCU\Software\VB and VBA Program Settings\RinCorelMacros\AutoAlignJobPresets`. Setiap nama adalah satu nilai dengan format berversi; penyimpanannya terpisah dari `AutoAlignDefaults/ProfilesV1`. Export/import preset dan penyimpanan geometri zona belum termasuk fitur ini.

## Pemasangan dari Source

Target penggunaan project adalah **CorelDRAW 2024 64-bit dengan VBA**. Kompatibilitas dengan versi lain perlu diperiksa tersendiri.

Repository menyediakan code-behind UserForm, Class Module, dan Standard Module. Paket GMS siap pakai serta layout UserForm `.frm/.frx` tidak disertakan, sehingga pemasangan memerlukan project VBA dan kontrol form yang disiapkan secara manual.

1. Buat atau gunakan project GMS AutoAlign di editor VBA CorelDRAW.
2. Siapkan Class Module untuk setiap file pada `src/classes/`, dengan properti `(Name)` sama persis seperti nama file tanpa ekstensi. File `.cls` berisi source headerless; masukkan isinya ke Class Module bernama sesuai, bukan mengandalkan metadata import class.
3. Masukkan seluruh Standard Module pada `src/modules/` dengan nama sesuai file. `AADebug` tetap dibutuhkan walaupun keluaran debug dinonaktifkan.
4. Buat UserForm **AutoAlignMenu**, **DefaultSettings**, **CustomMarginSettings**, **InspectionResults**, **ProductionSummary**, dan **JobPresets**. Pasang code-behind dari `src/forms/` pada form yang namanya cocok, lalu siapkan kontrol berikut.
5. Lakukan **Debug → Compile** pada project VBA dan selesaikan referensi atau kontrol yang belum cocok sebelum menjalankan form.

### AutoAlignMenu

| Tipe kontrol | `(Name)` | Kegunaan/pengaturan |
| --- | --- | --- |
| ListBox | `lbxObjects` | Daftar sumber/container. Kode mengatur pilihan Ctrl/Shift. |
| CommandButton | `cmdSetDesign`, `cmdSetCutLine` | Mendaftarkan selection sebagai Design atau Cut Line. |
| CommandButton | `cmdRemove`, `cmdClear` | Menghapus baris terpilih atau mengosongkan daftar. |
| Label | `lblCutPartner` | Caption **Pasangan Cut Line**. |
| ComboBox | `cmbCutPartner` | Style `fmStyleDropDownList`; RowSource/ControlSource kosong. Pilihan dan kolom ID diisi kode. |
| CommandButton | `cmdSelectPair` | Caption **Pilih Pasangan**. |
| Label | `lblRotation` | Caption **Rotasi Design**; beri ruang untuk teks **Rotasi Design (Campuran)**. |
| ComboBox | `cmbRotation` | Style `fmStyleDropDownList`; pilihan diisi kode. RowSource/ControlSource kosong. |
| TextBox | `txbQuantity` | Quantity per Design. |
| OptionButton | `optKissA`, `optDieA`, `optCustom` | Model area. |
| OptionButton | `optMinimum`, `optMedium`, `optMaximum` | Mode algoritma; Maximum dinonaktifkan kode. |
| TextBox | `txbAreaWidth`, `txbAreaHeight` | Ukuran area Custom. |
| TextBox | `txbGapHorizontal`, `txbGapVertical` | Gap H/V dalam mm. |
| CommandButton | `cmdDefaultSetting` | Membuka DefaultSettings. |
| CommandButton | `cmdCustomMarginSettings` | Membuka editor margin dan zona Custom. |
| CommandButton | `cmdProcess`, `cmdOptimize`, `cmdCheckLayout` | Process, Optimize, dan **Periksa Susunan**. |
| CommandButton | `cmdProductionSummary` | Caption **Ringkasan Hasil**; membuka catatan Process terakhir yang berhasil. |
| CommandButton | `cmdJobPresets` | Caption **Preset Pekerjaan**; menyimpan/menerapkan konfigurasi bernama. |
| CommandButton | `cmdClose` | Close/Cancel sesuai status operasi. |
| Frame | `frmProgress` | Area bar progres. |
| Label | `lblProgressFill` | Isi bar; letakkan di dalam `frmProgress`. |
| Label | `lblProgressStage`, `lblProgressDetail`, `lblProgressTime` | Tahap, detail, dan durasi. |

Tombol Default Settings, margin Custom, Ringkasan Hasil, dan Preset Pekerjaan menampilkan hint saat kursor diarahkan ke tombol. Hint Preset Pekerjaan juga menampilkan nama preset default yang sedang diterapkan.

Kode mengatur GroupName model/mode, daftar ComboBox, multi-selection, dan placeholder. Gunakan tipe serta nama kontrol persis seperti source; ukuran dan tata letak visual dapat disesuaikan.

### DefaultSettings

| Tipe kontrol | `(Name)` | Kegunaan |
| --- | --- | --- |
| OptionButton | `optKissASet`, `optDieASet` | Memilih profil default yang diedit. |
| TextBox | `txbDefaultAreaWidth`, `txbDefaultAreaHeight` | Area W/H dalam mm. |
| TextBox | `txbDefaultGapHorizontal`, `txbDefaultGapVertical` | Gap H/V dalam mm. |
| CommandButton | `cmdSave`, `cmdCancel` | Save kedua draft profil / Cancel tanpa menyimpan. |

### CustomMarginSettings

| Tipe kontrol | `(Name)` | Kegunaan |
| --- | --- | --- |
| TextBox | `txbMarginLeft`, `txbMarginRight`, `txbMarginTop`, `txbMarginBottom` | Empat margin dalam mm; tambahkan Label penjelas. |
| ListBox | `lbxForbiddenAreas` | Daftar zona, sumber, X/Y dan W/H; RowSource kosong. Beri lebar untuk ringkasan. |
| TextBox | `txbZoneX`, `txbZoneY`, `txbZoneWidth`, `txbZoneHeight` | Rectangle manual dalam mm; tambahkan Label penjelas. |
| CommandButton | `cmdAddRectangle`, `cmdAddFromSelection` | Tambah Rectangle / Ambil Seleksi. |
| CommandButton | `cmdRemoveZone`, `cmdClearZones` | Hapus Zona / Kosongkan Zona. |
| CheckBox | `chkIncludeDesign` | Caption **Design juga menghindari zona**; dua keadaan, awal dicentang oleh kode. |
| CommandButton | `cmdApply`, `cmdCancel` | Terapkan / Batal. |

Tambahkan Class Module headerless **AACustomArea**. Pembaruan fitur area mencakup class `AASettings`, `AAPresenter`, `AAProcessEngine`, `AAPlanner`, `AAMediumPlanner`, `AAModel`, `AAPlacement`, `AARotationVariant`, `AAResultBuilder`, `AAOptimizeEngine`, `AAOptimizeLayout`, `AAInspectionEngine`; module `AADebug`; serta kode form `AutoAlignMenu` dan `CustomMarginSettings`. Pertahankan class/form lain yang sudah dipasang. Tidak diperlukan macro publik tambahan atau resource `.frm/.frx` pengganti.

### InspectionResults

| Tipe kontrol | `(Name)` | Kegunaan/pengaturan |
| --- | --- | --- |
| Label | `lblSummary` | Ringkasan; sediakan ruang untuk beberapa baris. |
| ListBox | `lbxIssues` | Lima kolom yang diatur kode: kategori, area, ID pertama, ID kedua, dan detail. |
| TextBox | `txbDetail` | `MultiLine=True`, `Locked=True`; ringkasan, detail temuan, dan acuan pemeriksaan. |
| CommandButton | `cmdSelectIssue` | Caption **Pilih Objek**. |
| CommandButton | `cmdClose` | Caption **Tutup**. |

Form ini dibuka oleh `cmdCheckLayout`; tidak memerlukan macro publik tersendiri. Pastikan `AAInspectionEngine`, `AAInspectionReport`, dan `AAInspectionIssue` tersedia sebagai Class Module.

### ProductionSummary

| Tipe kontrol | `(Name)` | Kegunaan/pengaturan |
| --- | --- | --- |
| Label | `lblSummary` | Total Design, Cut Line fisik dan area; sediakan beberapa baris. |
| ListBox | `lbxProduction` | Empat kolom: Area, Design, Cut Line, Jumlah Design. RowSource kosong; kolom diisi kode. |
| TextBox | `txbDetails` | `MultiLine=True`, `Locked=True`, scrollbar vertikal; ringkasan dan detail. |
| CommandButton | `cmdCopy` | Caption **Salin Ringkasan**. |
| CommandButton | `cmdClose` | Caption **Tutup**. |

Tambahkan empat Label judul kolom di atas ListBox; nama Label bebas. Lebar kolom yang diatur kode adalah 45/85/85/85 pt.

Tambahkan Class Module headerless **AAProductionReport** dengan `(Name)` yang sama. Perbarui `AAProcessEngine`, `AAResultBuilder`, `AAPresenter`, serta kode form `AutoAlignMenu` dan `ProductionSummary`. Nama TextBox yang dipakai adalah **txbDetails**. Clipboard memakai `MSForms.DataObject` dari Microsoft Forms yang digunakan UserForm. Tidak diperlukan macro publik tambahan.

### JobPresets

| Tipe kontrol | `(Name)` | Kegunaan/pengaturan |
| --- | --- | --- |
| ListBox | `lbxPresets` | Pilihan tunggal; RowSource kosong. Baris pertama adalah pilihan Tanpa preset. |
| TextBox | `txbPresetName` | Nama preset; panjang maksimal 80 karakter. Tambahkan Label **Nama preset**. |
| TextBox | `txbDetails` | `MultiLine=True`, `Locked=True`, scrollbar vertikal. |
| CommandButton | `cmdSave` | Caption **Simpan Pengaturan Saat Ini**. |
| CommandButton | `cmdApply`, `cmdDelete`, `cmdClose` | Caption **Terapkan**, **Hapus**, **Tutup**. |

Tambahkan dua Class Module headerless **AAJobPreset** dan **AAJobPresetStore** dengan `(Name)` yang sama. Perbarui `AADefaultPreferences`, `AASettings`, `AAPresenter`, `AAProcessEngine`, `AAOptimizeEngine`, `AAInspectionEngine`, serta kode form `AutoAlignMenu` dan `JobPresets`. `AADefaultPreferences.ReadStoredNumber` dipakai bersama untuk membaca nilai numerik tersimpan. Tidak diperlukan macro publik baru atau perubahan resource form lain.

## Integrasi dan Source

| Lokasi | Peran |
| --- | --- |
| `src/forms/AutoAlignMenu.vba` | Pendaftaran objek, pilihan model/mode, input gap, Quantity, Process, dan callback MacroRunner. |
| `src/classes/AAPresenter.cls`, `AAObjectStore.cls`, `AASettings.cls` | Penghubung form, daftar objek, dan validasi pengaturan. |
| `src/classes/AADefaultPreferences.cls`, `src/forms/DefaultSettings.vba` | Preferensi Area/Gap per model, draft editor dan penyimpanan registry. |
| `src/classes/AAProcessEngine.cls`, `AAResultBuilder.cls` | Menjalankan proses, menangani pasangan/kelompok, dan membangun hasil. |
| `src/classes/AAProductionReport.cls`, `src/forms/ProductionSummary.vba` | Catatan hasil Process, hitungan Design/Cut Line fisik, dan salinan teks ringkasan. |
| `src/classes/AAJobPreset.cls`, `AAJobPresetStore.cls`, `src/forms/JobPresets.vba` | Snapshot pengaturan pekerjaan, penyimpanan bernama di registry dan dialog pengelolaan preset. |
| `src/classes/AAInspectionEngine.cls`, `AAInspectionReport.cls`, `AAInspectionIssue.cls`, `src/forms/InspectionResults.vba` | Pemeriksaan susunan terdaftar, laporan temuan, dan pemilihan objek terkait. |
| `src/classes/AAOptimizeEngine.cls`, `AAOptimizeMinimum.cls`, `AAOptimizeLayout.cls` | Optimize Minimum/Medium, snapshot kelompok, validasi, dan penerapan translasi. Import sebagai Class Module dengan nama yang sesuai. |
| `src/classes/AAPlanner.cls`, `AAMediumPlanner.cls`, `AAQuantityPlanner.cls` | Perencanaan Minimum, Medium, dan kuota kelompok. |
| `src/classes/AALayoutPlan.cls`, `AAPlacement.cls`, `AARotationVariant.cls`, `AAGroupBucket.cls` | Data rencana, posisi dan orientasi kandidat, serta pengelompokan hasil. |
| `src/classes/AAModel.cls`, `AAGeometry.cls`, `AAContour.cls`, `AAOccupiedIndex.cls` | Acuan bentuk, kontur, rotasi, gap, dan pemeriksaan tetangga. |
| `src/modules/AAAlgorithm.bas` | Konstanta serta aturan layer dan Z-Order. |
| `src/modules/AADebug.bas` | Debug, profiling, dan diagnostik memori. |
| `src/modules/AAGapBenchmark.bas` | Benchmark diagnostik dua rumus jarak segmen melalui Immediate Window. Import sebagai Standard Module bernama `AAGapBenchmark`. |
| `src/modules/AAProgress.bas` | Indikator tahap/detail/durasi, pembaruan bar, serta penguncian sementara jendela dokumen saat Process. |
| `src/modules/MRTargetBridge.bas` | Membuka form melalui `OpenMacro` dan menghubungkan callback MacroRunner. |
| [Changelog.log](Changelog.log) | Riwayat perubahan. |

Integrasi MacroRunner mendukung pembukaan/penutupan form dan pencatatan durasi Process melalui `MRTargetBridge.OpenMacro`, `MRBindRunner`, dan `MRDetachRunner`. Tidak diperlukan reference langsung ke project VBA MacroRunner. Form dengan nama yang sama harus ditutup sebelum dibuka melalui antrean. Pengaturan serta penyusunan dijalankan melalui form AutoAlign.

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

Pada pencarian kuota Medium, `Quota reject cache` mencatat pemeriksaan posisi gagal yang dipakai ulang (`hits`), pencarian tanpa kecocokan (`misses`), dan penggantian slot cache (`replacements`). Counter ini membantu membandingkan biaya trial pada kasus yang sama.

`edge batches` mencatat pemakaian query indeks untuk kelompok segmen: `queries` adalah query gabungan, `prechecks` pemeriksaan awal per kelompok, dan `rechecks` verifikasi penolakan melalui query per segmen. `fallback cells/spread/dense` membedakan batas luas sel, pelebaran query, dan kepadatan kandidat; `fallback` juga mencakup kelompok yang terlalu pendek. `source edges` menghitung segmen sumber pada jalur tanpa fallback, termasuk pemeriksaan awal. `offered` adalah jumlah kandidat yang disediakan untuk pemeriksaan batch, bukan jumlah pasangan yang benar-benar dihitung jaraknya.

Untuk membandingkan biaya fungsi jarak, import `AAGapBenchmark.bas`, lalu jalankan `AADebugBenchmarkGap` saat Process tidak berjalan. Helper memakai 4096 pasangan segmen sintetis yang tetap, memeriksa kesamaan hasil, melakukan pemanasan, lalu menjalankan delapan putaran secara default dengan urutan A/B bergantian. A memakai rumus empat akar, B rumus satu akar. `median paired ratio B/A` di bawah 1 berarti B lebih cepat pada corpus tersebut; angka ini tidak mewakili percepatan Process atau pengujian kurva native.

Benchmark tidak memerlukan dokumen atau seleksi. Hasil tampil di Immediate Window dan mengikuti file debug jika aktif. Contoh penyimpanan log:

```vb
AADebugOn
AADebugFileOn Environ$("TEMP") & "\AutoAlign-GapBenchmark.txt"
AADebugBenchmarkGap
AADebugOff
```

## Troubleshooting

| Kondisi | Yang perlu diperiksa |
| --- | --- |
| Design belum dipasangkan atau tampil **Pasangan hilang /ID** | Pilih Design, lalu tetapkan ulang `cmbCutPartner`. Untuk standalone, pilih **Tanpa Cut Line** dan hapus semua baris Cut Line. Remove tidak mengalihkan pasangan otomatis. |
| Ingin berbagi beberapa Cut Line ke beberapa kelompok Design | Gunakan satu Cut Line bersama untuk seluruh Design, atau satu Cut Line berbeda per Design. Campuran kelompok berbagi Cut Line belum didukung. |
| Nilai Quantity/rotasi tampak kosong saat memilih banyak Design | Pilihan dapat berbeda antarbaris. Memilih nilai baru menerapkannya ke semua Design terpilih; Cut Line diabaikan. |
| Jumlah hasil Medium berkurang setelah rotasi dibatasi | Kandidat orientasi yang boleh dipakai lebih sedikit. Pada Cut Line bersama, periksa irisan preset seluruh Design aktif. |
| Hasil tidak sesuai gap yang diinginkan | Isi Gap H/V manual dalam mm. Kolom kosong memakai default; pada rectangle bersudut runcing preset, default gap tetap 0 mm. Posisi diagonal memakai nilai H/V terbesar. |
| ID sumber tidak ditemukan setelah Process | Undo hasil Process sekali sebelum mencoba Process lain dengan sumber yang sama, atau kosongkan daftar dan daftarkan sumber/container yang memang ingin dipakai. |
| Optimize menolak susunan | Pastikan input berupa container, struktur anggota sesuai, pasangan unik, dan layer dapat diedit. Medium memerlukan susunan awal valid terhadap Area/Gap/mark. |
| Optimize lolos, tetapi Periksa Susunan menemukan batas area/mark | Optimize berpusat pada bounding box container; pemeriksaan memakai grid halaman Process. Periksa posisi container, ukuran area, model, dan halaman sumber. |
| Laporan **Pemeriksaan belum selesai** | Periksa detail: pembatalan, geometri/ID gagal dibaca, daftar duplikat, atau pasangan ambigu. Status ini tidak menyatakan susunan bebas masalah. |
| Cancel belum langsung berhenti | Permintaan ditangani pada checkpoint. Tunggu operasi Corel yang sedang berjalan dan proses pemulihan selesai. |
| Compile gagal atau hasil pemeriksaan tidak bisa dibuka | Cocokkan `(Name)` Class Module, UserForm, dan seluruh kontrol; pastikan `InspectionResults` serta class pemeriksaan tersedia. |

## Status dan Batasan

- Process menyediakan mode **Minimum** dan **Medium**. `optMaximum` dinonaktifkan. Optimize bersifat eksperimental; periksa hasil pemadatannya sebelum digunakan.
- Perencanaan dibatasi hingga **5000 objek per sumber**; satu Cut Line bersama juga mempunyai batas jumlah terkait.
- Sumber harus mempunyai ukuran yang valid dan kurva yang dapat dibaca. Pada group yang perlu dibaca geometrinya, anggota tanpa kurva dapat menyebabkan validasi gagal.
- Daftar mempertahankan ID sumber. Jika sumber dihapus, dipindah, atau berubah struktur, pulihkan sumber atau daftarkan ulang sebelum Process.
- Preset Quantity, batas rotasi, pasangan Cut Line, dan daftar Objects berlaku selama instance form aktif. Hanya default Area/Gap KissA dan DieA yang disimpan ke registry.
- Pemeriksaan hanya mencakup objek terdaftar dan memakai validator kontur AA. Laporan tanpa temuan tidak membuktikan semua objek halaman aman atau memberi pengukuran gap eksak yang independen.
- Optimize dan Periksa Susunan membaca struktur container tertentu; keduanya tidak menelusuri group manual dengan kedalaman sembarang.
- Hasil bergantung pada bentuk, jumlah sumber, model area, gap, Quantity, dan batas pencarian. Pencarian heuristik tidak menjamin nesting optimum.

## Lisensi dan Masukan

Project ini menggunakan **MIT License**; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
