# CK Hardware Scanner

Lightweight, portable hardware diagnostics for Catur Komputer technicians.

## Fitur

- Informasi komputer, BIOS, Windows, CPU, RAM, disk, baterai, GPU, dan adapter jaringan.
- Detail RAM per modul dari `Win32_PhysicalMemory`: slot/bank, kapasitas, manufacturer, part number, serial, SMBIOS memory type, speed, configured clock, voltage, form factor, rank, dan field ketersediaan bila dilaporkan firmware.
- Detail CPU, motherboard, BIOS, TPM/Secure Boot, dan Windows system: core, threads, clock saat scan, cache, virtualization, firmware/boot metadata, build, boot time, dan uptime bila dilaporkan Windows.
- Report TXT dan HTML juga menampilkan detail CPU, RAM per modul, motherboard, BIOS, Windows, Secure Boot, dan TPM; JSON mempertahankan struktur teknis yang sama.
- Phase 3C menampilkan physical disk, partition/volume, ruang terpakai/kosong bila dapat dibaca, serta detail battery dari WMI. SMART health, temperatur storage, power-on hours, dan statistik lifetime ditandai unavailable karena tidak disediakan oleh sumber WMI standar yang digunakan.
- Phase 3D menampilkan detail GPU dan network adapter dari WMI/configuration Windows, termasuk dedicated/shared memory distinction, PNP/status, IP, gateway, DNS, DHCP, link speed, dan klasifikasi physical/virtual bila dilaporkan sumber.
- Phase 3E menyediakan provider sensor LHM terisolasi sebagai child process. PowerShell 5.1 memakai helper Framework, sedangkan PowerShell 7 memakai helper .NET 10; keduanya mengirim JSON dengan schema yang sama dan watchdog 30 detik.
- Deteksi perangkat Plug-and-Play dengan `ConfigManagerErrorCode`.
- Status ringkas `OK`, `WARNING`, `ERROR`, dan `N/A`.
- Kode masalah PnP dipertahankan dan diklasifikasikan konservatif; kode 22 dan 24 menghasilkan `WARNING`, bukan diagnosis kerusakan fisik.
- Output terminal dan laporan TXT, HTML, atau JSON opsional dari hasil scan yang sama.
- Modular dan disiapkan untuk ekspor HTML/JSON pada fase berikutnya.

## Persyaratan

- Windows PowerShell 5.1 pada Windows 10 atau Windows 11 untuk jalur helper Framework.
- PowerShell 7 menggunakan helper LHM target .NET 10 dan memerlukan .NET 10 runtime yang kompatibel; runtime tersebut tidak dipasang atau diunduh otomatis.
- Tidak memerlukan Python, Node.js, atau modul PowerShell tambahan.
- Pemeriksaan dasar dan helper sensor tidak meminta Administrator secara otomatis.

## Cara menjalankan

```powershell
.\scan.ps1
```

Untuk menyimpan laporan setelah menjalankan scanner:

```powershell
.\scan.ps1 -SaveReport -OutputDirectory .\reports
```

Format laporan:

```powershell
.\scan.ps1 -SaveReport -OutputFormat HTML -OutputDirectory .\reports
.\scan.ps1 -SaveReport -OutputFormat JSON -OutputDirectory .\reports
```

Untuk menyamarkan serial number dan MAC address pada output/laporan:

```powershell
.\scan.ps1 -SaveReport -RedactSensitive
```

HTML dapat dibuka offline dan memiliki tombol Print yang tidak ikut tercetak. JSON berisi struktur data, bukan HTML atau teks gabungan.

Nama laporan memakai hostname dan timestamp; serial number tidak digunakan sebagai nama file. Scanner tidak mengirim data ke internet.

## Batasan pemeriksaan

Data bergantung pada dukungan firmware dan driver Windows. `N/A` berarti data tidak tersedia, bukan bukti kerusakan. Status keseluruhan `ERROR` hanya untuk kegagalan scanner kritis; masalah PnP menghasilkan `WARNING` dan tidak membuktikan kerusakan fisik. Status disk memakai status yang dilaporkan Windows dan bukan pengganti pemeriksaan forensik atau stress test. Scanner tidak menjalankan stress test dan tidak mengubah konfigurasi Windows. Kompatibilitas Windows 10/11 belum diverifikasi di kedua lingkungan pada repository ini.

Detail RAM memakai data WMI/SMBIOS yang tersedia. Generasi DDR tidak ditebak dari speed, voltage tidak diturunkan dari generasi DDR, dan `RatedSpeed`/data channel ditampilkan `N/A` bila firmware tidak melaporkannya. Nilai SMBIOS yang tidak dikenal tetap disimpan sebagai raw value dengan status `Unknown`.

Detail CPU dan system juga mengikuti data WMI/Windows yang benar-benar tersedia. `BaseClock`, cache, socket, TDP, instruction set, chipset, TPM, Secure Boot, dan UEFI mode dapat menjadi `N/A` atau `Not available` bila sumber tidak melaporkannya; current clock tidak digunakan sebagai base clock.

Windows `DisplayVersion` dan `UBR` dibaca dari registry Windows bila tersedia. Keduanya tidak ditebak dari build number; jika tidak tersedia, report menggunakan `N/A`.

Serial number RAM, disk, dan battery dipertahankan untuk diagnosis internal. Opsi masking lama tetap dapat menyamarkan serial system dan MAC address, tetapi tidak menghapus serial hardware diagnosis.

GPU temperature, GPU power, storage temperature, dan system power ditampilkan `Unavailable` bila tidak ada sensor Windows standar yang tervalidasi. Tidak ada driver atau aplikasi sensor pihak ketiga yang dipasang otomatis.

Jika provider sensor tidak tersedia, scanner tetap berjalan dan menampilkan `Unavailable`/`N/A`; kondisi tersebut bukan indikasi kerusakan hardware.

Provider LHM dibundel pada `runtimes/windows-x64/powershell51` dan `runtimes/windows-x64/powershell7`. Asset LHM v0.9.6 dan SHA-256 dicatat pada `manifest.json`; lisensi MPL-2.0 dan third-party notices disertakan di masing-masing folder. Helper .NET 10 memerlukan runtime .NET 10 yang kompatibel pada komputer pengguna dan tidak memasang runtime secara otomatis. Jika helper, dependency, runtime, atau sensor tidak tersedia, scan hardware utama tetap dilanjutkan.

Nilai sensor yang tidak diberikan LHM tetap `null`/`Unavailable`; nilai nol yang diberikan LHM dipertahankan sebagai nol. Unit hanya dipetakan untuk tipe yang terverifikasi seperti Load `%`, Temperature `C`, Clock `MHz`, Power `W`, Fan `RPM`, Voltage `V`, dan Current `A`.

## Pengujian

Jika Pester tersedia, dapat digunakan sesuai konfigurasi lokal. Tanpa Pester, jalankan test runner bawaan:

```powershell
.\tests\Run-Tests.ps1
```

Test runner hanya memakai PowerShell bawaan dan mock ringan; beberapa tes hardware membaca data mesin saat ini.
