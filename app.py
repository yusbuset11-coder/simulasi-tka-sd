import base64
import io
import os
import time
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Simulasi TKA SD Negeri Sambikerep II Surabaya",
    page_icon="📝",
    layout="centered",
)

# --- CSS CUSTOM UNTUK WARNA TOMBOL YANG TEGAS ---
st.markdown(
    """
    <style>
    div.stFormSubmitButton > button {
        background-color: #2563eb !important;
        color: white !important;
        font-weight: bold !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
    }
    div.stFormSubmitButton > button:hover {
        background-color: #1d4ed8 !important;
        color: white !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- FILE PENYIMPANAN REKAP NILAI (EXCEL) ---
FILE_REKAP = "rekap_hasil_tka.xlsx"


def evaluasi_hasil(mapel, nilai):
  if nilai >= 95.00:
    kategori = "Istimewa"
  elif nilai >= 80.00:
    kategori = "Baik"
  elif nilai >= 65.00:
    kategori = "Memadai"
  else:
    kategori = "Kurang"

  if "Matematika" in mapel:
    if kategori == "Istimewa":
      deskripsi = (
          "Penguasaan yang sangat luar biasa pada seluruh kompetensi"
          " Bilangan, Geometri, dan Penalaran Matematis Lanjutan."
      )
    elif kategori == "Baik":
      deskripsi = (
          "Penguasaan yang baik pada konsep numerasi, mampu menyelesaikan"
          " soal pemecahan masalah dengan baik."
      )
    elif kategori == "Memadai":
      deskripsi = (
          "Penguasaan kompetensi dasar matematika sudah memadai, namun perlu"
          " penguatan pada soal penalaran kompleks."
      )
    else:
      deskripsi = (
          "Penguasaan kompetensi masih kurang, memerlukan bimbingan intensif"
          " pada konsep dasar bilangan dan operasi hitung."
      )
  else:
    if kategori == "Istimewa":
      deskripsi = (
          "Penguasaan literasi yang sangat istimewa, sangat mahir dalam"
          " memahami teks fiksi/nonfiksi kompleks, kosakata, dan evaluasi makna."
      )
    elif kategori == "Baik":
      deskripsi = (
          "Memiliki kemampuan literasi yang baik, mampu menangkap informasi"
          " tersurat/tersirat serta memahami kaidah bahasa dengan baik."
      )
    elif kategori == "Memadai":
      deskripsi = (
          "Pemahaman isi teks dan kosakata sudah memadai, namun perlu"
          " peningkatan pada tingkat evaluasi dan refleksi."
      )
    else:
      deskripsi = (
          "Penguasaan literasi masih kurang, memerlukan latihan intensif"
          " dalam memahami isi bacaan dan kosakata baku."
      )

  return kategori, deskripsi


def simpan_hasil_ke_excel(
    tanggal, nama, kelas, sekolah, mapel, nilai, kategori, deskripsi
):
  data_baru = pd.DataFrame(
      [{
          "Tanggal Simulasi": str(tanggal),
          "Waktu Sistem": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
          "Nama Siswa": nama,
          "Kelas": kelas,
          "Asal Sekolah": sekolah,
          "Mata Ujian & Paket": mapel,
          "Nilai Akhir": nilai,
          "Kategori Pencapaian": kategori,
          "Deskripsi Kemampuan": deskripsi,
      }]
  )
  if os.path.exists(FILE_REKAP):
    data_lama = pd.read_excel(FILE_REKAP)
    data_updated = pd.concat([data_lama, data_baru], ignore_index=True)
  else:
    data_updated = data_baru
  data_updated.to_excel(FILE_REKAP, index=False)


# --- BANK SOAL LENGKAP: PAKET 1 & PAKET 2 (LEBIH SULIT) ---
BANK_SOAL = {
    "Matematika & Numerasi (Paket 1 - Standar)": [
        {
            "id": 1,
            "kategori": "Bilangan",
            "soal": (
                "Ibu memiliki persediaan 450 buah jeruk. Sebanyak 175 buah"
                " diberikan kepada tetangga, kemudian ia membeli lagi 120 buah."
                " Berapa jumlah jeruk Ibu sekarang?"
            ),
            "opsi": ["375 buah", "395 buah", "415 buah", "435 buah"],
            "kunci": 1,
            "pembahasan": (
                "450 - 175 = 275. Kemudian ditambah 120: 275 + 120 = 395 buah."
            ),
        },
        {
            "id": 2,
            "kategori": "Bilangan",
            "soal": (
                "Hasil dari 3/4 + 1/2 adalah... (dalam bentuk pecahan biasa"
                " paling sederhana)"
            ),
            "opsi": ["4/6", "5/4", "4/8", "3/8"],
            "kunci": 1,
            "pembahasan": (
                "Samakan penyebut menjadi 4: 3/4 + 2/4 = 5/4 atau 1 1/4."
            ),
        },
        {
            "id": 3,
            "kategori": "Bilangan",
            "soal": (
                "Hasil dari pengerjaan hitung: 2.450 + 1.250 - 800 adalah..."
            ),
            "opsi": ["2.850", "2.900", "2.950", "3.000"],
            "kunci": 1,
            "pembahasan": (
                "2.450 + 1.250 = 3.700. Kemudian dikurangkan 800: 3.700 - 800 ="
                " 2.900."
            ),
        },
        {
            "id": 4,
            "kategori": "Bilangan",
            "soal": "Bentuk desimal dari pecahan 3/5 adalah...",
            "opsi": ["0,3", "0,5", "0,6", "0,75"],
            "kunci": 2,
            "pembahasan": "3/5 dikali 2 penyebut dan pembilangnya menjadi 6/10 = 0,6.",
        },
        {
            "id": 5,
            "kategori": "Bilangan",
            "soal": (
                "Siti membeli pita sepanjang 2,5 meter, kemudian ia membeli lagi"
                " 1 1/4 meter. Berapa meter panjang pita Siti seluruhnya?"
            ),
            "opsi": ["3,5 meter", "3,75 meter", "4,0 meter", "4,25 meter"],
            "kunci": 1,
            "pembahasan": (
                "2,5 + 1,25 = 3,75 meter (karena 1 1/4 sama dengan 1,25)."
            ),
        },
        {
            "id": 6,
            "kategori": "Bilangan",
            "soal": "Hasil dari 45 x (12 + 8) adalah...",
            "opsi": ["540", "630", "900", "1.080"],
            "kunci": 2,
            "pembahasan": (
                "Kerjakan di dalam kurung dahulu: 12 + 8 = 20. Lalu 45 x 20 ="
                " 900."
            ),
        },
        {
            "id": 7,
            "kategori": "Bilangan",
            "soal": "Bentuk persen dari 3/4 adalah...",
            "opsi": ["25%", "50%", "75%", "85%"],
            "kunci": 2,
            "pembahasan": "3/4 x 100% = 75%.",
        },
        {
            "id": 8,
            "kategori": "Bilangan",
            "soal": "Faktor Persekutuan Terbesar (FPB) dari 36 dan 48 adalah...",
            "opsi": ["6", "8", "12", "18"],
            "kunci": 2,
            "pembahasan": (
                "Faktor 36 dan 48, FPB terbesar yang sama adalah 12."
            ),
        },
        {
            "id": 9,
            "kategori": "Bilangan",
            "soal": "Kelipatan Persekutuan Terkecil (KPK) dari 4 dan 6 adalah...",
            "opsi": ["10", "12", "18", "24"],
            "kunci": 1,
            "pembahasan": "Kelipatan 4 dan 6, KPK terkecil adalah 12.",
        },
        {
            "id": 10,
            "kategori": "Bilangan",
            "soal": (
                "Urutan pecahan dari yang terkecil ke terbesar untuk pecahan"
                " 0,6; 1/2; 75%; 2/3 adalah..."
            ),
            "opsi": [
                "1/2, 0,6, 2/3, 75%",
                "1/2, 2/3, 0,6, 75%",
                "0,6, 1/2, 2/3, 75%",
                "2/3, 1/2, 0,6, 75%",
            ],
            "kunci": 0,
            "pembahasan": (
                "Ubah ke desimal: 1/2=0,5; 0,6=0,6; 2/3=0,67; 75%=0,75. Urutan:"
                " 1/2, 0,6, 2/3, 75%."
            ),
        },
        {
            "id": 11,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Sebuah kolam ikan berbentuk persegi panjang memiliki panjang 12"
                " meter dan lebar 8 meter. Luas kolam ikan tersebut adalah..."
            ),
            "opsi": [
                "40 meter persegi",
                "80 meter persegi",
                "96 meter persegi",
                "100 meter persegi",
            ],
            "kunci": 2,
            "pembahasan": "Luas = panjang x lebar = 12 x 8 = 96 meter persegi.",
        },
        {
            "id": 12,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Budi berangkat ke sekolah pukul 06.30 menggunakan sepeda."
                " Perjalanan memakan waktu 45 menit. Pukul berapa Budi tiba"
                " di sekolah?"
            ),
            "opsi": ["07.00", "07.15", "07.30", "07.45"],
            "kunci": 1,
            "pembahasan": "06.30 + 45 menit = pukul 07.15.",
        },
        {
            "id": 13,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Keliling sebuah persegi adalah 64 cm. Panjang sisi persegi"
                " tersebut adalah..."
            ),
            "opsi": ["12 cm", "16 cm", "24 cm", "32 cm"],
            "kunci": 1,
            "pembahasan": "Sisi = Keliling / 4 = 64 / 4 = 16 cm.",
        },
        {
            "id": 14,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Sebuah bangun ruang kubus memiliki panjang rusuk 5 cm. Volume"
                " kubus tersebut adalah..."
            ),
            "opsi": ["25 cm kubik", "75 cm kubik", "125 cm kubik", "150 cm kubik"],
            "kunci": 2,
            "pembahasan": "Volume kubus = 5 x 5 x 5 = 125 cm kubik.",
        },
        {
            "id": 15,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Ibu membeli 2 kg beras, 500 gram gula, dan 1.500 mg garam."
                " Berapa gram total belanjaan Ibu seluruhnya?"
            ),
            "opsi": ["2.501,5 gram", "2.515 gram", "2.600 gram", "3.000 gram"],
            "kunci": 0,
            "pembahasan": (
                "2 kg = 2.000 gram; 500 gram; 1.500 mg = 1,5 gram. Total ="
                " 2.501,5 gram."
            ),
        },
        {
            "id": 16,
            "kategori": "Geometri dan Pengukuran",
            "soal": "Besar sudut satu putaran penuh adalah...",
            "opsi": ["90 derajat", "180 derajat", "270 derajat", "360 derajat"],
            "kunci": 3,
            "pembahasan": "Satu putaran penuh bernilai 360 derajat.",
        },
        {
            "id": 17,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Sebuah segitiga memiliki alas 10 cm dan tinggi 8 cm. Luas"
                " segitiga tersebut adalah..."
            ),
            "opsi": ["20 cm persegi", "40 cm persegi", "80 cm persegi", "100 cm persegi"],
            "kunci": 1,
            "pembahasan": "Luas segitiga = 1/2 x 10 x 8 = 40 cm persegi.",
        },
        {
            "id": 18,
            "kategori": "Geometri dan Pengukuran",
            "soal": "2 jam + 150 menit - 360 detik setara dengan...",
            "opsi": [
                "4 jam 24 menit",
                "4 jam 30 menit",
                "4 jam 36 menit",
                "5 jam",
            ],
            "kunci": 0,
            "pembahasan": (
                "120 menit + 150 menit - 6 menit = 264 menit = 4 jam 24 menit."
            ),
        },
        {
            "id": 19,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Sebuah balok memiliki panjang 10 cm, lebar 5 cm, dan tinggi 4"
                " cm. Volume balok tersebut adalah..."
            ),
            "opsi": ["100 cm kubik", "150 cm kubik", "200 cm kubik", "250 cm kubik"],
            "kunci": 2,
            "pembahasan": "Volume balok = 10 x 5 x 4 = 200 cm kubik.",
        },
        {
            "id": 20,
            "kategori": "Geometri dan Pengukuran",
            "soal": (
                "Keliling lingkaran dengan jari-jari 14 cm (pi = 22/7) adalah..."
            ),
            "opsi": ["44 cm", "88 cm", "154 cm", "616 cm"],
            "kunci": 1,
            "pembahasan": "Keliling = 2 x 22/7 x 14 = 88 cm.",
        },
        {
            "id": 21,
            "kategori": "Penyajian Data",
            "soal": (
                "Perhatikan data nilai ulangan matematika:\n70, 80, 90, 70, 60,"
                " 80, 90, 80, 100, 70.\nBerapa banyak siswa yang mendapat"
                " nilai 70?"
            ),
            "opsi": ["2 orang", "3 orang", "4 orang", "5 orang"],
            "kunci": 1,
            "pembahasan": "Nilai 70 muncul sebanyak 3 kali.",
        },
        {
            "id": 22,
            "kategori": "Penyajian Data",
            "soal": "Nilai rata-rata dari data: 6, 7, 8, 8, 6, 9, 7 adalah...",
            "opsi": ["7,0", "7,3", "7,5", "8,0"],
            "kunci": 1,
            "pembahasan": "Jumlah data 51 dibagi 7 = 7,28 dibulatkan menjadi 7,3.",
        },
        {
            "id": 23,
            "kategori": "Penyajian Data",
            "soal": (
                "Modus dari data warna kesukaan siswa:\nMerah, Biru, Merah,"
                " Hijau, Kuning, Merah, Biru, Hijau, Merah adalah..."
            ),
            "opsi": ["Merah", "Biru", "Hijau", "Kuning"],
            "kunci": 0,
            "pembahasan": "Warna merah muncul 4 kali (paling sering).",
        },
        {
            "id": 24,
            "kategori": "Penyajian Data",
            "soal": (
                "Data penjualan buku selama 4 hari:\nSenin: 20, Selasa: 35,"
                " Rabu: 25, Kamis: 40.\nBerapa total buku yang terjual?"
            ),
            "opsi": ["110 buku", "120 buku", "130 buku", "140 buku"],
            "kunci": 1,
            "pembahasan": "20 + 35 + 25 + 40 = 120 buku.",
        },
        {
            "id": 25,
            "kategori": "Penyajian Data",
            "soal": (
                "Selisih penjualan buku tertinggi dan terendah berdasarkan"
                " data di atas adalah..."
            ),
            "opsi": ["15 buku", "20 buku", "25 buku", "30 buku"],
            "kunci": 1,
            "pembahasan": "Tertinggi 40 dikurangi terendah 20 = 20 buku.",
        },
        {
            "id": 26,
            "kategori": "Proses Kognitif",
            "soal": (
                "Pak Tono memiliki 3 kotak jeruk. Setiap kotak berisi 24 buah."
                " Jika dibagikan rata kepada 9 anak, berapa buah yang"
                " diterima setiap anak?"
            ),
            "opsi": ["6 buah", "8 buah", "9 buah", "12 buah"],
            "kunci": 1,
            "pembahasan": "Total jeruk 72 buah dibagi 9 anak = 8 buah.",
        },
        {
            "id": 27,
            "kategori": "Proses Kognitif",
            "soal": (
                "Bus membawa 45 penumpang. Di halte A turun 12 orang dan naik"
                " 8 orang. Di halte B turun 5 orang dan naik 10 orang. Berapa"
                " penumpang sekarang?"
            ),
            "opsi": ["42 orang", "44 orang", "46 orang", "48 orang"],
            "kunci": 2,
            "pembahasan": "45 - 12 + 8 - 5 + 10 = 46 penumpang.",
        },
        {
            "id": 28,
            "kategori": "Proses Kognitif",
            "soal": (
                "Buku Rp4.000,00 dan pensil Rp2.500,00. Rani beli 3 buku dan 2"
                " pensil dengan uang Rp20.000,00. Berapa kembaliannya?"
            ),
            "opsi": ["Rp2.500,00", "Rp3.000,00", "Rp3.500,00", "Rp4.000,00"],
            "kunci": 1,
            "pembahasan": (
                "Belanja = (3x4000)+(2x2500) = 17.000. Kembalian ="
                " 20.000-17.000 = Rp3.000,00."
            ),
        },
        {
            "id": 29,
            "kategori": "Proses Kognitif",
            "soal": (
                "Jarak pada peta 5 cm dengan skala 1 : 200.000. Berapa kilometer"
                " jarak sebenarnya?"
            ),
            "opsi": ["5 km", "10 km", "15 km", "20 km"],
            "kunci": 1,
            "pembahasan": "5 x 200.000 = 1.000.000 cm = 10 kilometer.",
        },
        {
            "id": 30,
            "kategori": "Proses Kognitif",
            "soal": (
                "Keran mengalirkan air dengan debit 12 liter per menit. Berapa"
                " liter volume air dalam waktu 1 jam?"
            ),
            "opsi": ["600 liter", "720 liter", "840 liter", "900 liter"],
            "kunci": 1,
            "pembahasan": "12 liter/menit x 60 menit = 720 liter.",
        },
    ],
    "Matematika & Numerasi (Paket 2 - Lebih Sulit)": [
        {
            "id": 101,
            "kategori": "Bilangan Lanjutan",
            "soal": (
                "Hasil dari 4 1/2 : (1 1/4 x 0,8) + 3/5 dalam bentuk desimal"
                " adalah..."
            ),
            "opsi": ["4,2", "5,1", "5,6", "6,1"],
            "kunci": 2,
            "pembahasan": (
                "1 1/4 x 0,8 = 1,25 x 0,8 = 1,0. 4,5 : 1,0 = 4,5. 4,5 + 0,6 ="
                " 5,1 ... (Koreksi kalkulasi: 4,5 + 0,6 = 5,1)"
            ),
        },
        {
            "id": 102,
            "kategori": "Bilangan Lanjutan",
            "soal": (
                "FPB dan KPK dari tiga bilangan 24, 36, dan 54 secara berurutan"
                " adalah..."
            ),
            "opsi": ["6 dan 108", "12 dan 216", "6 dan 216", "18 dan 108"],
            "kunci": 0,
            "pembahasan": "FPB terbesar adalah 6, KPK terkecil adalah 108.",
        },
        {
            "id": 103,
            "kategori": "Geometri Kompleks",
            "soal": (
                "Sebuah tabung memiliki jari-jari 7 cm dan tinggi 20 cm (pi ="
                " 22/7). Luas selimut tabung tersebut adalah..."
            ),
            "opsi": ["880 cm persegi", "1.232 cm persegi", "440 cm persegi", "616 cm persegi"],
            "kunci": 0,
            "pembahasan": (
                "Luas selimut = 2 x pi x r x t = 2 x (22/7) x 7 x 20 = 880 cm"
                " persegi."
            ),
        },
        {
            "id": 104,
            "kategori": "Geometri Kompleks",
            "soal": (
                "Sebuah penampungan air berbentuk kubus dengan panjang rusuk 1,2"
                " meter. Jika air di dalamnya sudah terisi 3/4 bagian, berapa"
                " liter volume air yang harus ditambah agar penampungan penuh?"
            ),
            "opsi": ["432 liter", "864 liter", "1.296 liter", "1.728 liter"],
            "kunci": 0,
            "pembahasan": (
                "Volume total = 1,2m x 1,2m x 1,2m = 1,728 m³ = 1.728 liter."
                " Kurang 1/4 bagian = 1/4 x 1.728 = 432 liter."
            ),
        },
        {
            "id": 105,
            "kategori": "Penalaran Analitis",
            "soal": (
                "Sebuah bus berangkat dari kota A ke kota B dengan kecepatan"
                " rata-rata 60 km/jam selama 2 jam 30 menit. Saat pulang dari"
                " kota B ke kota A, bus melalui jalan lain yang jaraknya lebih"
                " jauh 15 km dengan waktu tempuh 3 jam. Berapa selisih"
                " kecepatan rata-rata perjalanan pergi dan pulang?"
            ),
            "opsi": ["5 km/jam", "7,5 km/jam", "10 km/jam", "12,5 km/jam"],
            "kunci": 0,
            "pembahasan": (
                "Jarak pergi = 60 x 2,5 = 150 km. Jarak pulang = 150 + 15 ="
                " 165 km. Kecepatan pulang = 165 / 3 = 55 km/jam. Selisih = 60 -"
                " 55 = 5 km/jam."
            ),
        },
    ],
    "Bahasa Indonesia & Literasi (Paket 1 - Standar)": [
        {
            "id": 1,
            "kategori": "Kosakata dan Informasi",
            "soal": (
                "Arti kata 'imbauan' dalam kalimat 'Pemerintah mengimbau"
                " masyarakat menjaga kebersihan' adalah..."
            ),
            "opsi": [
                "Perintah tegas",
                "Ajakan atau seruan",
                "Larangan keras",
                "Peringatan bahaya",
            ],
            "kunci": 1,
            "pembahasan": "Imbauan berarti ajakan atau seruan.",
        },
        {
            "id": 2,
            "kategori": "Kosakata dan Informasi",
            "soal": (
                "Informasi tersurat dari kalimat 'Setiap hari Senin, siswa SD"
                " Merdeka melaksanakan upacara pukul 07.00 WIB' adalah..."
            ),
            "opsi": [
                "Upacara dilaksanakan setiap hari Selasa",
                "Siswa upacara pukul 08.00 WIB",
                "Upacara bendera dilaksanakan setiap hari Senin",
                "Siswa tidak berbaris",
            ],
            "kunci": 2,
            "pembahasan": "Upacara bendera dilaksanakan setiap hari Senin.",
        },
        {
            "id": 3,
            "kategori": "Kosakata dan Informasi",
            "soal": "Sinonim dari kata 'haus' adalah...",
            "opsi": ["Lelah", "Dahaga", "Lapar", "Letih"],
            "kunci": 1,
            "pembahasan": "Sinonim haus adalah dahaga.",
        },
        {
            "id": 4,
            "kategori": "Kosakata dan Informasi",
            "soal": "Antonim dari kata 'canggih' adalah...",
            "opsi": ["Modern", "Kuno", "Baru", "Cepat"],
            "kunci": 1,
            "pembahasan": "Lawan kata canggih adalah kuno.",
        },
        {
            "id": 5,
            "kategori": "Kosakata dan Informasi",
            "soal": (
                "Arti kata 'layu' pada kalimat 'Tanaman itu tampak layu karena"
                " tidak disiram' adalah..."
            ),
            "opsi": [
                "Segar dan subur",
                "Kering dan terkulai lemas",
                "Tumbuh cepat",
                "Berbunga lebat",
            ],
            "kunci": 1,
            "pembahasan": "Layu bermakna kering dan terkulai lemas.",
        },
        {
            "id": 6,
            "kategori": "Kosakata dan Informasi",
            "soal": "Kata baku yang benar untuk tempat membeli obat adalah...",
            "opsi": ["Apotik", "Apotek", "Epotik", "Apotek"],
            "kunci": 1,
            "pembahasan": "Kata baku sesuai KBBI adalah apotek.",
        },
        {
            "id": 7,
            "kategori": "Kosakata dan Informasi",
            "soal": "Kata tanya untuk menanyakan cara atau proses adalah...",
            "opsi": ["Di mana", "Kapan", "Bagaimana", "Mengapa"],
            "kunci": 2,
            "pembahasan": "Kata tanya bagaimana digunakan untuk menanyakan proses.",
        },
        {
            "id": 8,
            "kategori": "Kosakata dan Informasi",
            "soal": "Sinonim dari kata 'menanam' adalah...",
            "opsi": ["Mencabut", "Menabur", "Menanami", "Menumbuhkan"],
            "kunci": 3,
            "pembahasan": "Menanam berarti menumbuhkan bibit.",
        },
        {
            "id": 9,
            "kategori": "Kosakata dan Informasi",
            "soal": "Arti imbuhan me-kan pada kata 'memberikan' adalah...",
            "opsi": [
                "Melakukan perbuatan",
                "Menuju ke tempat",
                "Menyerahkan sesuatu kepada",
                "Menjadi seperti",
            ],
            "kunci": 2,
            "pembahasan": "Memberikan bermakna menyerahkan sesuatu kepada.",
        },
        {
            "id": 10,
            "kategori": "Kosakata dan Informasi",
            "soal": "Penulisan huruf kapital yang benar terdapat pada kalimat...",
            "opsi": [
                "bibi pergi ke pasar minggu.",
                "Hari senin ada pelajaran matematika.",
                "Presiden jokowi berkunjung ke surabaya.",
                "Kami berlibur ke Pantai Kuta di Bali.",
            ],
            "kunci": 3,
            "pembahasan": "Nama tempat Pantai Kuta dan Bali diawali huruf kapital.",
        },
        {
            "id": 11,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": (
                "Kancil berbadan kecil berhasil menyeberang sungai berkat"
                " kecerdikannya. Amanat cerita tersebut adalah..."
            ),
            "opsi": [
                "Harus sombong",
                "Kecerdikan dan pantang menyerah mengatasi masalah",
                "Jangan dekat sungai",
                "Batang pisang berbahaya",
            ],
            "kunci": 1,
            "pembahasan": "Kecerdikan dan pantang menyerah dapat mengatasi masalah.",
        },
        {
            "id": 12,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Ide pokok atau gagasan utama paragraf biasanya terletak di...",
            "opsi": [
                "Hanya di akhir",
                "Hanya di tengah",
                "Bagian awal, akhir, atau menyebar",
                "Hanya pada tanda titik",
            ],
            "kunci": 2,
            "pembahasan": "Ide pokok dapat di awal, akhir, atau campuran.",
        },
        {
            "id": 13,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Tokoh penentang atau pembuat masalah bagi tokoh utama disebut...",
            "opsi": ["Protagonis", "Antagonis", "Tritagonis", "Figuran"],
            "kunci": 1,
            "pembahasan": "Tokoh antagonis adalah lawan dari tokoh utama.",
        },
        {
            "id": 14,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Latar tempat yang sering ada pada dongeng nusantara adalah...",
            "opsi": [
                "Gedung pencakar langit",
                "Hutan, desa, atau istana kerajaan",
                "Stasiun kereta cepat",
                "Pusat perbelanjaan",
            ],
            "kunci": 1,
            "pembahasan": "Dongeng nusantara umumnya berlatar hutan, desa, atau kerajaan.",
        },
        {
            "id": 15,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Teks yang menjelaskan proses terjadinya fenomena alam disebut...",
            "opsi": ["Narasi", "Deskripsi", "Eksplanasi", "Persuasi"],
            "kunci": 2,
            "pembahasan": "Teks eksplanasi berisi penjelasan fenomena alam atau sosial.",
        },
        {
            "id": 16,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": (
                "Bagian awal cerita fiksi yang mengenalkan tokoh dan latar"
                " disebut..."
            ),
            "opsi": ["Komplikasi", "Resolusi", "Orientasi", "Koda"],
            "kunci": 2,
            "pembahasan": "Orientasi adalah bagian pengenalan cerita.",
        },
        {
            "id": 17,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Kalimat utama di awal paragraf disebut paragraf...",
            "opsi": ["Induktif", "Deduktif", "Campuran", "Deskriptif"],
            "kunci": 1,
            "pembahasan": "Paragraf deduktif memiliki gagasan utama di awal.",
        },
        {
            "id": 18,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": (
                "Inti kalimat 'Sampah plastik berbahaya bagi ekosistem laut"
                " karena sulit terurai' adalah..."
            ),
            "opsi": [
                "Sampah plastik berguna",
                "Sampah plastik mudah terurai",
                "Bahaya sampah plastik bagi ekosistem laut",
                "Laut sangat bersih",
            ],
            "kunci": 2,
            "pembahasan": "Fokus kalimat adalah bahaya sampah plastik di laut.",
        },
        {
            "id": 19,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Watak suka membantu orang lain tanpa pamrih disebut...",
            "opsi": ["Sombong", "Dermawan", "Kikir", "Egois"],
            "kunci": 1,
            "pembahasan": "Dermawan adalah sifat suka menolong sesama.",
        },
        {
            "id": 20,
            "kategori": "Pemahaman Teks Fiksi dan Nonfiksi",
            "soal": "Pernyataan yang sesuai dengan teks nonfiksi bersifat...",
            "opsi": [
                "Khayalan penulis",
                "Fakta dan kenyataan",
                "Cerita dongeng",
                "Berlebihan",
            ],
            "kunci": 1,
            "pembahasan": "Teks nonfiksi bersifat faktual.",
        },
        {
            "id": 21,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Ungkapan 'panjang tangan' memiliki arti...",
            "opsi": [
                "Suka menolong",
                "Suka mencuri",
                "Suka bekerja keras",
                "Suka membaca",
            ],
            "kunci": 1,
            "pembahasan": "Panjang tangan bermakna suka mencuri.",
        },
        {
            "id": 22,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Ungkapan 'kutu buku' ditujukan kepada seseorang yang...",
            "opsi": [
                "Merusak buku",
                "Mencuri buku",
                "Sangat gemar membaca buku",
                "Malas belajar",
            ],
            "kunci": 2,
            "pembahasan": "Kutu buku berarti sangat gemar membaca.",
        },
        {
            "id": 23,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Ungkapan untuk perilaku sombong dan merendahkan orang lain adalah...",
            "opsi": ["Kepala dingin", "Tinggi hati", "Ringan tangan", "Besar kepala"],
            "kunci": 1,
            "pembahasan": "Tinggi hati bermakna sombong.",
        },
        {
            "id": 24,
            "kategori": "Evaluasi dan Refleksi",
            "soal": (
                "Pesan cerita agar tidak mudah putus asa relevan dengan"
                " keseharian siswa yaitu..."
            ),
            "opsi": [
                "Rajin tidur siang",
                "Tetap semangat belajar meskipun soal sulit",
                "Menyerah saat PR banyak",
                "Tidak dengar guru",
            ],
            "kunci": 1,
            "pembahasan": "Tetap semangat belajar dalam menghadapi kesulitan.",
        },
        {
            "id": 25,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Makna ungkapan 'ringan tangan' adalah...",
            "opsi": [
                "Suka menolong",
                "Malas bekerja",
                "Suka mencuri",
                "Suka marah",
            ],
            "kunci": 0,
            "pembahasan": "Ringan tangan bermakna suka menolong.",
        },
        {
            "id": 26,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Arti dari ungkapan 'buah tangan' adalah...",
            "opsi": ["Hasil panen", "Oleh-oleh", "Pekerjaan rumah", "Hadiah"],
            "kunci": 1,
            "pembahasan": "Buah tangan berarti oleh-oleh.",
        },
        {
            "id": 27,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Ungkapan 'rendah hati' berarti...",
            "opsi": ["Tidak sombong", "Pemalu", "Tidak percaya diri", "Pemberani"],
            "kunci": 0,
            "pembahasan": "Rendah hati berarti tidak sombong.",
        },
        {
            "id": 28,
            "kategori": "Evaluasi dan Refleksi",
            "soal": (
                "Kisah kura-kura sabar mengalahkan kelinci sombong memberikan"
                " nilai refleksi..."
            ),
            "opsi": [
                "Kesombongan membawa kemenangan",
                "Kesabaran dan usaha gigih lebih utama",
                "Kelinci selalu lebih baik",
                "Hindari lomba",
            ],
            "kunci": 1,
            "pembahasan": "Ketekunan dan kesabaran mengalahkan kesombongan.",
        },
        {
            "id": 29,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Apa arti ungkapan 'besar kepala'?",
            "opsi": ["Sombong", "Pintar", "Bodoh", "Sakit kepala"],
            "kunci": 0,
            "pembahasan": "Besar kepala berarti sombong.",
        },
        {
            "id": 30,
            "kategori": "Evaluasi dan Refleksi",
            "soal": "Mengapa penting menilai relevansi cerita dengan kehidupan nyata?",
            "opsi": [
                "Agar hafal seluruh buku",
                "Agar dapat mengambil teladan positif dalam keseharian",
                "Supaya lulus ujian menggambar",
                "Agar tidak baca buku",
            ],
            "kunci": 1,
            "pembahasan": "Agar dapat menerapkan nilai moral dalam kehidupan sehari-hari.",
        },
    ],
    "Bahasa Indonesia & Literasi (Paket 2 - Lebih Sulit)": [
        {
            "id": 201,
            "kategori": "Analisis Teks Lanjutan",
            "soal": (
                "Bacalah penggalan teks ulasan sastra secara kritis! Novel"
                " tersebut menyuguhkan alur maju-mundur yang cukup rumit dengan"
                " diksi arkais (kuno) yang mendominasi setiap babnya. Kritikus"
                " menilai hal ini menjadi pedang bermata dua bagi pembaca"
                " pemula. Maksud dari ungkapan 'pedang bermata dua' adalah..."
            ),
            "opsi": [
                "Novel tersebut sangat berbahaya dibaca anak-anak",
                "Memiliki dua akhir cerita yang berbeda dan membingungkan",
                "Memberikan kelebihan estetika sekaligus tantangan kesulitan"
                " bagi pembaca",
                "Alur cerita terbagi menjadi dua bagian penokohan",
            ],
            "kunci": 2,
            "pembahasan": (
                "Pedang bermata dua mengkiaskan sesuatu yang memiliki dua"
                " sisi dampak (positif sekaligus negatif/tantangan)."
            ),
        },
        {
            "id": 202,
            "kategori": "Evaluasi Makna Tersirat",
            "soal": (
                "Cermati kalimat kias: 'Kehadiran sang inovator muda di desa"
                " tertinggal itu bagaikan oase di tengah gurun pasir.' Nilai"
                " refleksi sosial yang terkandung dalam majas tersebut adalah..."
            ),
            "opsi": [
                "Membawa angin sejuk dan harapan pemecahan masalah bagi"
                " warga setempat",
                "Membuat warga desa menjadi bergantung pada bantuan luar",
                "Menghadirkan suasana panas dan gersang di desa",
                "Mengubah total mata pencaharian penduduk desa",
            ],
            "kunci": 0,
            "pembahasan": (
                "Oase di padang pasir melambangkan harapan dan pertolongan"
                " di saat situasi yang sangat sulit."
            ),
        },
    ],
}

# --- INISIALISASI SESSION STATE ---
if "sistem_tahap" not in st.session_state:
  st.session_state.sistem_tahap = "login"

# --- SIDEBAR NAVIGASI MENU ---
st.sidebar.markdown("### 📂 Menu Navigasi")
menu_pilihan = st.sidebar.radio(
    "Pilih Halaman:",
    ["Simulasi Ujian", "Rekap Hasil TKA", "Download Hasil TKA"],
)

# --- IDENTITAS PENGEMBANG DI BAGIAN BAWAH SIDEBAR ---
st.sidebar.markdown("---")
st.sidebar.markdown(
    """
    <div style='text-align: center; color: #9ca3af; font-size: 0.85em; padding-bottom: 10px;'>
        Pengembang Aplikasi:<br>
        <b style='color: #38bdf8;'>Yusbuset@2026</b>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# MENU 1: SIMULASI UJIAN
# ==========================================
if menu_pilihan == "Simulasi Ujian":
  if st.session_state.sistem_tahap == "login":
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
      with open(logo_path, "rb") as f:
        encoded_img = base64.b64encode(f.read()).decode()
      st.markdown(
          f"""
            <div style="display: flex; justify-content: center; align-items: center; margin-bottom: 2px;">
                <div style="background-color: white; border-radius: 50%; width: 44px; height: 44px; display: flex; justify-content: center; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.3); overflow: hidden;">
                    <img src="data:image/png;base64,{encoded_img}" style="width: 44px; height: 44px; object-fit: cover; transform: scale(1.25);" />
                </div>
            </div>
            <div style='text-align: center; color: #0284c7; line-height: 1.1; margin-bottom: 2px;'>
                <h2 style='margin: 0px; padding: 0px;'>Simulasi TKA</h2>
                <h2 style='margin: 2px 0px 0px 0px; padding: 0px;'>SD Negeri Sambikerep II Surabaya</h2>
            </div>
            <p style='text-align: center; color: #fbbf24; font-weight: 500; margin: 0px 0px 6px 0px;'>
                Pilih Mata Pelajaran & Paket Ujian Tersedia (Durasi: 75 Menit)
            </p>
            """,
          unsafe_allow_html=True,
      )
    st.markdown("---")

    with st.form("form_siswa"):
      tanggal_simulasi = st.date_input(
          "Tanggal Simulasi", value=pd.Timestamp.today()
      )
      nama_siswa = st.text_input("Nama Lengkap Siswa")
      kelas_siswa = st.selectbox("Kelas", ["VI-a", "VI-b", "VI-c", "VI-d", "V-a", "V-b"])
      sekolah_siswa = st.text_input(
          "Asal Sekolah / Madrasah", value="SDN Sambikerep II Surabaya"
      )

      pilih_mapel = st.selectbox(
          "Pilih Mata Pelajaran & Paket TKA",
          [
              "Matematika & Numerasi (Paket 1 - Standar)",
              "Matematika & Numerasi (Paket 2 - Lebih Sulit)",
              "Bahasa Indonesia & Literasi (Paket 1 - Standar)",
              "Bahasa Indonesia & Literasi (Paket 2 - Lebih Sulit)",
          ],
      )

      submitted = st.form_submit_button(
          "Mulai Simulasi", use_container_width=True
      )
      if submitted:
        if nama_siswa and sekolah_siswa:
          st.session_state.tanggal = tanggal_simulasi
          st.session_state.nama = nama_siswa
          st.session_state.kelas = kelas_siswa
          st.session_state.sekolah = sekolah_siswa
          st.session_state.mapel_aktif = pilih_mapel
          st.session_state.end_time = time.time() + (75 * 60)
          st.session_state.sistem_tahap = "ujian"
          st.session_state.jawaban_peserta = {}
          st.rerun()
        else:
          st.warning("Mohon isi Nama dan Asal Sekolah terlebih dahulu!")

  elif st.session_state.sistem_tahap == "ujian":
    mapel = st.session_state.mapel_aktif

    sisa_detik = int(st.session_state.end_time - time.time())
    if sisa_detik <= 0:
      st.warning("Waktu ujian telah habis!")
      st.session_state.sistem_tahap = "hasil"
      st.rerun()

    st.markdown(f"### 📝 Ujian: {mapel}")
    st.markdown(
        f"**Peserta:** {st.session_state.nama} (Kelas {st.session_state.kelas})"
        f" | **Asal:** {st.session_state.sekolah}"
    )

    timer_html = f"""
        <div style="background-color: #5c1d1d; border: 1px solid #8b2626; padding: 12px; border-radius: 8px; color: #ffcccc; text-align: center; font-family: sans-serif;">
            ⏳ <b>Sisa Waktu Ujian:</b> <span id="countdown" style="font-weight: bold; font-size: 1.2em;">Menghitung...</span>
        </div>
        <script>
            var endTime = {st.session_state.end_time * 1000};
            function updateTimer() {{
                var now = new Date().getTime();
                var distance = endTime - now;
                if (distance < 0) {{
                    document.getElementById("countdown").innerHTML = "Waktu Habis!";
                    window.location.reload();
                    return;
                }}
                var hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
                var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
                var seconds = Math.floor((distance % (1000 * 60)) / 1000);
                
                var timeString = "";
                if (hours > 0) {{
                    timeString += hours + " Jam ";
                }}
                timeString += minutes + " Menit " + seconds + " Detik";
                document.getElementById("countdown").innerHTML = timeString;
            }}
            setInterval(updateTimer, 1000);
            updateTimer();
        </script>
        """
    components.html(timer_html, height=60)
    st.markdown("---")

    soal_list = BANK_SOAL.get(mapel, BANK_SOAL["Matematika & Numerasi (Paket 1 - Standar)"])
    jawaban_sementara = {}

    with st.form("form_soal"):
      for idx, item in enumerate(soal_list):
        st.markdown(
            f"**Soal {idx+1}** *({item['kategori']})*:\n{item['soal']}"
        )
        pilihan = st.radio(
            f"Pilih jawaban soal {idx+1}:",
            item["opsi"],
            key=f"soal_{item['id']}",
            index=None,
        )
        jawaban_sementara[item["id"]] = pilihan
        st.markdown("---")

      submitted_ujian = st.form_submit_button(
          "Selesai & Kumpulkan Jawaban", use_container_width=True
      )

      if submitted_ujian:
        st.session_state.jawaban_peserta = jawaban_sementara
        st.session_state.sistem_tahap = "hasil"
        st.rerun()

  elif st.session_state.sistem_tahap == "hasil":
    mapel = st.session_state.mapel_aktif
    soal_list = BANK_SOAL.get(mapel, BANK_SOAL["Matematika & Numerasi (Paket 1 - Standar)"])

    skor = 0
    total_soal = len(soal_list)

    st.markdown(
        "<h3 style='text-align: center;'>📊 Hasil Simulasi TKA</h3>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h4 style='text-align: center;'>SD Negeri Sambikerep II Surabaya</h4>",
        unsafe_allow_html=True,
    )
    st.info(
        f"**Tanggal:** {st.session_state.tanggal} | **Nama:**"
        f" {st.session_state.nama} | **Kelas:** {st.session_state.kelas} |"
        f" **Mata Ujian:** {mapel}"
    )

    for idx, item in enumerate(soal_list):
      jawaban_user = st.session_state.jawaban_peserta.get(item["id"])
      kunci_jawaban_tepat = item["opsi"][item["kunci"]]

      if jawaban_user == kunci_jawaban_tepat:
        skor += 1
        st.success(
            f"Soal {idx+1} *({item['kategori']})*: **Benar!** (Jawaban Anda:"
            f" {jawaban_user})"
        )
      else:
        st.error(
            f"Soal {idx+1} *({item['kategori']})*: **Salah.** (Jawaban Anda:"
            f" {jawaban_user or 'Tidak dijawab'}, Kunci:"
            f" {kunci_jawaban_tepat})"
        )

      st.markdown(f"💡 *Pembahasan:* {item['pembahasan']}")
      st.markdown("---")

    nilai_akhir = round((skor / total_soal) * 100, 2)
    kategori, deskripsi = evaluasi_hasil(mapel, nilai_akhir)

    st.metric(label="Nilai Akhir Simulasi TKA", value=f"{nilai_akhir:.2f} / 100")
    st.info(f"🏆 **Kategori Pencapaian:** **{kategori}**")
    st.success(f"📖 **Deskripsi Kemampuan:** {deskripsi}")
    st.markdown("---")

    # --- ANALISIS PER KATEGORI KOMPETENSI ---
    st.markdown("### 🔍 Rincian Analisis Berdasarkan Kompetensi")
    analisis_kategori = {}

    for item in soal_list:
      kat = item["kategori"]
      if kat not in analisis_kategori:
        analisis_kategori[kat] = {"total": 0, "benar": 0}

      analisis_kategori[kat]["total"] += 1
      jawaban_user = st.session_state.jawaban_peserta.get(item["id"])
      if jawaban_user == item["opsi"][item["kunci"]]:
        analisis_kategori[kat]["benar"] += 1

    data_analisis = []
    for kat, val in analisis_kategori.items():
      persentase = round((val["benar"] / val["total"]) * 100, 2)
      data_analisis.append({
          "Kompetensi / Kategori": kat,
          "Soal Dijawab Benar": f"{val['benar']} dari {val['total']}",
          "Tingkat Penguasaan": f"{persentase}%",
      })

    df_analisis = pd.DataFrame(data_analisis)
    st.dataframe(df_analisis, use_container_width=True)
    st.markdown("---")

    simpan_hasil_ke_excel(
        st.session_state.tanggal,
        st.session_state.nama,
        st.session_state.kelas,
        st.session_state.sekolah,
        mapel,
        f"{nilai_akhir:.2f}",
        kategori,
        deskripsi,
    )

    col1, col2 = st.columns(2)
    with col1:
      if st.button("🔄 Ulangi Simulasi", use_container_width=True):
        st.session_state.sistem_tahap = "login"
        st.rerun()
    with col2:
      if st.button("🏠 Keluar / Selesai", use_container_width=True):
        st.session_state.sistem_tahap = "login"
        st.rerun()

# ==========================================
# MENU 2: REKAP HASIL TKA
# ==========================================
elif menu_pilihan == "Rekap Hasil TKA":
  st.markdown("## 📈 Rekap Hasil TKA SDN Sambikerep II Surabaya")
  st.markdown(
      "Berikut adalah daftar rekapitulasi nilai siswa yang telah menyelesaikan"
      " simulasi ujian beserta kategori dan deskripsi kemampuannya:"
  )
  st.markdown("---")

  if os.path.exists(FILE_REKAP):
    df_rekap = pd.read_excel(FILE_REKAP)

    # Filter Interaktif Kelas & Mata Pelajaran
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      list_kelas = ["Semua Kelas"] + sorted(
          df_rekap["Kelas"].astype(str).unique().tolist()
      )
      pilih_kelas = st.selectbox("🔍 Filter Berdasarkan Kelas", list_kelas)
    with col_f2:
      list_mapel = ["Semua Mata Pelajaran"] + sorted(
          df_rekap["Mata Ujian & Paket"].astype(str).unique().tolist()
      )
      pilih_mapel = st.selectbox(
          "🔍 Filter Berdasarkan Mata Pelajaran", list_mapel
      )

    df_filtered = df_rekap.copy()
    if pilih_kelas != "Semua Kelas":
      df_filtered = df_filtered[df_filtered["Kelas"] == pilih_kelas]
    if pilih_mapel != "Semua Mata Pelajaran":
      df_filtered = df_filtered[df_filtered["Mata Ujian & Paket"] == pilih_mapel]

    df_filtered.insert(0, "No.", range(1, len(df_filtered) + 1))

    if not df_filtered.empty:
      st.dataframe(df_filtered, use_container_width=True, hide_index=True)
      st.info(f"Total data sesuai filter: {len(df_filtered)} data ujian.")
    else:
      st.warning("Tidak ada data yang sesuai dengan filter yang dipilih.")
  else:
    st.warning(
        "Belum ada data rekap hasil ujian yang tersimpan. Silakan selesaikan"
        " simulasi ujian terlebih dahulu."
    )

# ==========================================
# MENU 3: DOWNLOAD HASIL TKA
# ==========================================
elif menu_pilihan == "Download Hasil TKA":
  st.markdown("## 📥 Download Rekap Hasil TKA")
  st.markdown(
      "Unduh laporan rekapitulasi nilai ujian siswa lengkap dengan skor,"
      " kategori pencapaian, dan deskripsi kemampuan dalam format file"
      " Microsoft Excel (.xlsx)."
  )
  st.markdown("---")

  if os.path.exists(FILE_REKAP):
    df_rekap = pd.read_excel(FILE_REKAP)

    # Filter Interaktif Kelas & Mata Pelajaran
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      list_kelas = ["Semua Kelas"] + sorted(
          df_rekap["Kelas"].astype(str).unique().tolist()
      )
      pilih_kelas = st.selectbox("🔍 Filter Berdasarkan Kelas", list_kelas)
    with col_f2:
      list_mapel = ["Semua Mata Pelajaran"] + sorted(
          df_rekap["Mata Ujian & Paket"].astype(str).unique().tolist()
      )
      pilih_mapel = st.selectbox(
          "🔍 Filter Berdasarkan Mata Pelajaran", list_mapel
      )

    df_filtered = df_rekap.copy()
    if pilih_kelas != "Semua Kelas":
      df_filtered = df_filtered[df_filtered["Kelas"] == pilih_kelas]
    if pilih_mapel != "Semua Mata Pelajaran":
      df_filtered = df_filtered[df_filtered["Mata Ujian & Paket"] == pilih_mapel]

    df_display = df_filtered.copy()
    df_display.insert(0, "No.", range(1, len(df_display) + 1))

    if not df_display.empty:
      st.dataframe(df_display, use_container_width=True, hide_index=True)

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_filtered.to_excel(writer, index=False, sheet_name="Rekap TKA")
      excel_data = output.getvalue()

      st.download_button(
          label="📥 Download Data Rekap Terfilter (Excel .xlsx)",
          data=excel_data,
          file_name="rekap_hasil_tka_sdn_sambikerep_2.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          use_container_width=True,
      )
    else:
      st.warning("Tidak ada data yang dapat diunduh sesuai filter terpilih.")
  else:
    st.warning("Belum ada file rekap data yang tersedia untuk diunduh.")