# Mulmed AI

Mulmed AI adalah sistem AI berbasis Python yang mengubah slide presentasi PowerPoint secara otomatis saat presenter selesai mengucapkan kalimat terakhir dari slide tersebut. Transisi dilakukan secara real-time menggunakan pengenalan suara, dan divalidasi oleh model AI dari OpenAI (GPT-4) & Google Gemini.

##  Fitur:

* Deteksi audio presenter dengan **Whisper**

* Deteksi akhir kalimat menggunakan 2–3 kata terakhir

* Validasi transisi dengan Gemini

* Kontrol PowerPoint melalui COM API (Windows only)

* Transisi presisi tinggi dengan delay minimal (0.75s)

##  Persyaratan (Requirements):

### Sistem:

* Windows (karena menggunakan COM untuk PowerPoint)

* PowerPoint terinstall dan presentasi dalam mode slideshow

## Python Packages:

```
pip install torch transformers openai google-generativeai pyaudio SpeechRecognition pywin32
```

## Instalasi

1. Clone repositori ini:
```bash
git clone https://github.com/Georgearr/Mulmed-AI.git
cd Mulmed-AI
```

2. **(Opsional)** Buat dan aktifkan virtual environment:
```bash
python -m venv venv
venv\Scripts\activate  # untuk Windows
```

3. Install semua requirements:
```bash
pip install -r requirements.txt
```

4. Tambahkan API key sebagai environment variable:
```bash
set OPENAI_API_KEY="[Masukkan API Key OpenAI]"
set GOOGLE_API_KEY="[Masukkan API Key Google Gemini]"
```

Atau buat file `.env`:
```env
OPENAI_API_KEY=[Masukkan API Key OpenAI]
GOOGLE_API_KEY=[Masukkan API Key Google Gemini]
```

---

## Konfigurasi API
Pastikan Anda memiliki API key dari OpenAI dan Google:

```bash
export OPENAI_API_KEY="[Masukkan API Key OpenAI]"
export GOOGLE_API_KEY="[Masukkan API Key Google Gemini]"
```

> Anda bisa menambahkan ini ke dalam `.env` dan gunakan `python-dotenv` jika ingin otomatis.

---


## Cara Menjalankan

1. Buka PowerPoint dan aktifkan `Slide Show (F5).`

2. Jalankan script:
`python audio_slide_changer_precise_full.py`

3. Slide akan berganti otomatis saat akhir kalimat terdeteksi.

## Cara Kerja

1. Script mendengarkan suara secara terus-menerus menggunakan microphone. Transkripsi dilakukan oleh model **Whisper**.

2. **2–3 kata terakhir dari transkripsi** dibandingkan dengan **akhir teks slide**.

3. Jika cocok **(kemiripan > 85%)**, script akan mengganti slide.

4. Gemini akan mengecek apakah **5 kata pertama** setelah transisi sesuai dengan **awal teks slide** baru.

## Catatan Penting

* Pastikan presentasi aktif di PowerPoint dan berada di **urutan pertama**.

* Gunakan input device dengan **kualitas baik** untuk **hasil maksimal**.

* Gemini dan ChatGPT hanya digunakan **untuk verifikasi**, bukan keputusan utama.

## Lisensi

Proyek ini open-source dan dapat dimodifikasi sesuai kebutuhan. AI ini sebenarnya dibuat untuk kebutuhan Multimedia, tetapi AI ini juga bebas digunakan untuk edukasi, presentasi, demo AI, dan lain-lain.

## Kontak

Dikembangkan oleh **George A. T.**
Untuk pertanyaan, bisa email georgearrev@gmail.com, atau DM Instagram dengan akun [@george_arrev_turnip](https://www.instagram.com/george_arrev_turnip)

Terimakasih.

Copyright © 2025 - George A. T.
