# hr-letter-generator

[English](README.md) | **Bahasa Melayu** | [中文](README.zh.md)

Satu alat ringkas yang tidak memerlukan pengalaman, untuk menjana **surat peletakan jawatan**
profesional dalam bentuk **Word (.docx)** dan **PDF (.pdf)**, dalam **Bahasa Inggeris, Bahasa
Melayu, atau Bahasa Cina (Ringkas)**.

Anda hanya menjawab beberapa soalan, dan ia akan:

- menulis surat mengikut standard HR / perekrut berpengalaman, dalam bahasa pilihan anda,
- mengira **hari terakhir bekerja** anda secara automatik daripada tarikh notis + tempoh notis (dalam bulan),
- menyimpan kedua-dua fail ke dalam folder `generated_letters/`.

Semua tarikh menggunakan format **`yyyy-MM-dd`** (cth. `2026-06-28`).

---

## Cara menggunakan

### Windows
1. Klik dua kali **`run.cmd`**.
2. Jawab soalan di dalam tetingkap yang terbuka.
3. Surat yang siap akan dibuka di dalam folder `generated_letters`.

### macOS
1. Klik dua kali **`run.command`**.
   - Kali pertama sahaja: jika macOS menyekatnya, klik kanan → **Open** → **Open**.
   - Jika klik dua kali tidak berfungsi, buka Terminal dalam folder ini dan jalankan `chmod +x run.command run.sh`, kemudian cuba lagi.
2. Jawab soalan di dalam tetingkap Terminal.
3. Surat yang siap akan dibuka di dalam folder `generated_letters`.

> Pengguna Terminal (macOS/Linux) juga boleh menjalankan `./run.sh`.

Pada kali pertama dijalankan, pelancar akan memeriksa komputer anda dan **memasang apa-apa yang
tiada secara automatik**:
- **Python 3** — melalui `winget` pada Windows atau **Homebrew** pada macOS. Jika ia tidak dapat
  dipasang secara automatik, tetingkap akan menunjukkan langkah yang perlu anda ambil.
- Pakej Python yang diperlukan (`python-docx`, `fpdf2`) — dipasang ke dalam folder `.venv`
  setempat supaya tiada apa-apa lain pada komputer anda terjejas.

Kali pertama mengambil masa kira-kira seminit untuk pemasangan; kali seterusnya bermula serta-merta.

---

## Apa yang ditanya

- **Bahasa:** Inggeris (lalai), Bahasa Melayu, atau 中文 (简体) — keseluruhan surat ditulis dalam bahasa ini
- **Maklumat anda:** nama penuh, jawatan, jabatan, e-mel, telefon
- **Penerima:** nama, jawatan, syarikat / organisasi, alamat
- **Notis:** tarikh anda memberi notis (lalai hari ini) dan tempoh notis dalam bulan
- **Pilihan:** satu pencapaian utama untuk disebut dalam surat

Selepas itu, ia memaparkan **hari terakhir bekerja yang dikira** supaya anda boleh menerimanya atau
menaip tarikh lain, serta satu ringkasan untuk semakan sebelum menjana surat.

> Petua: taip nama, syarikat dan pencapaian anda dalam bahasa yang anda pilih — semuanya akan
> muncul dalam surat tepat seperti yang anda taip. (Soalan pada skrin tetap dalam Bahasa Inggeris.)

### Bagaimana hari terakhir bekerja dikira

`hari terakhir bekerja = tarikh notis + tempoh notis (bulan) − 1 hari`

Contohnya, memberi notis pada `2026-06-01` dengan tempoh notis `1` bulan memberikan hari terakhir
bekerja `2026-06-30`. Anda sentiasa boleh menggantikan tarikh yang dikira semasa menjawab soalan.
Memasukkan `0` bulan akan menghasilkan surat "berkuat kuasa serta-merta".

---

## Output

Fail disimpan dalam `generated_letters/`. Namanya menggunakan awalan mengikut bahasa pilihan
diikuti nama anda, cth.:

```
Resignation_Letter_Jane_Doe.docx             (Inggeris)
Surat_Peletakan_Jawatan_Ahmad_bin_Ali.docx   (Bahasa Melayu)
辞职信_张伟.pdf                                 (中文)
```

(Jika fail dengan nama yang sama sudah wujud, satu nombor ditambah supaya tiada fail ditimpa.)

---

## Nota

- Fail **Word** boleh disunting sepenuhnya dan menyokong sebarang aksara.
- Untuk **Inggeris / Bahasa Melayu**, PDF menggunakan fon terbina standard; simbol yang sangat luar
  biasa mungkin tidak muncul dalam PDF — gunakan fail Word untuk itu.
- Untuk **Cina**, PDF membenamkan fon Cina yang terdapat pada komputer anda (cth. Microsoft YaHei
  atau SimSun pada Windows, PingFang pada macOS) — fon ini disertakan pada hampir semua sistem.
  Dalam kes yang jarang berlaku di mana tiada fon dijumpai, fail Word (.docx) masih dihasilkan
  dengan surat Cina yang lengkap.
- Tiada apa-apa dihantar melalui internet kecuali pemasangan pakej pada kali pertama. Tiada
  perkhidmatan berbayar digunakan.

---

## Penyelesaian masalah

- **Windows – "Python has been installed… close this window and run again."** Pemasangan kali
  pertama memerlukan tetingkap baharu untuk mengesan Python. Tutup tetingkap dan klik dua kali
  `run.cmd` sekali lagi.
- **Windows – winget tidak tersedia.** Pasang Python 3 dari <https://www.python.org/downloads/>
  dan tandakan *Add python.exe to PATH*, kemudian jalankan `run.cmd` semula.
- **macOS – "cannot be opened because it is from an unidentified developer."** Klik kanan
  `run.command` → **Open** → **Open** (hanya perlu pada kali pertama).
- **macOS – klik dua kali tidak berfungsi.** Buka Terminal dalam folder ini dan jalankan
  `chmod +x run.command run.sh`, kemudian cuba lagi.
- **Tiada internet pada kali pertama.** Pemasangan pakej kali pertama memerlukan internet. Selepas
  berjaya kali pertama, alat ini berfungsi di luar talian.
- **PDF Cina tiada.** Jika komputer tiada fon Cina, PDF dilangkau tetapi fail Word (`.docx`) masih
  mengandungi surat Cina yang lengkap (buka fail itu dan *Save As PDF* dari Word jika anda perlukan PDF).

---

## Fail dalam projek ini

| Fail | Apakah ia |
|------|------------|
| `run.cmd` | Pelancar Windows (klik dua kali) |
| `run.command` | Pelancar macOS (klik dua kali dalam Finder) |
| `run.sh` | Pelancar macOS / Linux (terminal) |
| `generate_letter.py` | program yang menanya soalan dan menulis surat |
| `requirements.txt` | pakej Python yang dipasang secara automatik |
| `.gitignore`, `.gitattributes` | penyelenggaraan repo (mengabaikan `.venv/` & `generated_letters/`; mengekalkan penghujung baris skrip dengan betul) |
| `generated_letters/` | tempat surat siap disimpan (dicipta pada kali pertama; tidak dimasukkan ke dalam repo) |

---

## Keperluan (dikendalikan secara automatik)

- Python 3.9+
- `python-docx`, `fpdf2` (lihat `requirements.txt`)
- Untuk PDF Cina: fon CJK yang sedia ada pada sistem (standard pada Windows & macOS)
