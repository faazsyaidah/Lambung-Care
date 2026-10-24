# ============================================================
# SISTEM PENDUKUNG DIAGNOSIS PENYAKIT LAMBUNG
# Metode: IVIFS + IIFWAA + Jarak Hamming
# Platform: Streamlit | Ekspor: Excel multi-pasien
#
# Jalankan:
#   pip install streamlit pandas numpy openpyxl
#   streamlit run app.py
#
# Catatan: prototipe penelitian, bukan pengganti diagnosis dokter.
# Validasi rumus dan data acuan dengan jurnal/dokter sebelum digunakan.
# ============================================================

from datetime import datetime
import io
import json
import math
import re
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="LambungCare | Sistem Pendukung Diagnosis",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded",
)

GEJALA = [
    "Mual",
    "Muntah",
    "Mudah kenyang",
    "Nyeri ulu hati",
    "Dada terasa terbakar (heartburn)",
    "Rasa penuh setelah makan",
    "Penurunan berat badan",
]

PENYAKIT = ["Gastritis", "Dispepsia", "GERD", "Gastroparesis", "Kanker lambung"]

# Skor acuan Dokter 3 sesuai kode sumber yang diberikan pengguna.
DATA_ACUAN = {
    "Gastritis":      [6, 6, 1, 9, 1, 1, 1],
    "Dispepsia":      [6, 6, 8, 6, 1, 8, 1],
    "GERD":           [8, 8, 6, 8, 8, 6, 1],
    "Gastroparesis":  [6, 6, 8, 3, 3, 8, 6],
    "Kanker lambung": [8, 8, 8, 6, 6, 8, 8],
}
LEBAR_INTERVAL = 0.10
DELTA = 1.6
BOBOT = [1.0] * len(GEJALA)

# --------------------------- Tampilan ---------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: linear-gradient(180deg, #f4f8ff 0%, #f8fafc 45%, #ffffff 100%); }
.block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1450px; }
.hero {
  padding: 1.65rem 1.8rem; border-radius: 22px;
  background: linear-gradient(120deg, #123b67 0%, #176b87 55%, #2c8c9d 100%);
  color: white; margin-bottom: 1.25rem;
  box-shadow: 0 12px 32px rgba(18,59,103,.15);
}
.hero h1 { color: white; font-size: 2rem; margin: 0 0 .35rem 0; font-weight: 800; }
.hero p { color: #e3f4ff; margin: 0; font-size: 1rem; }
.section-card {
  background: white; border: 1px solid #e4ebf4; border-radius: 18px;
  padding: 1.1rem 1.25rem; box-shadow: 0 4px 18px rgba(15, 40, 75, .045);
  margin-bottom: 1rem;
}
.small-muted { color: #62748a; font-size: .9rem; }
div[data-testid="stMetric"] {
  background: white; border: 1px solid #e4ebf4; padding: 1rem;
  border-radius: 16px; box-shadow: 0 3px 12px rgba(15,40,75,.04);
}
.stButton > button, .stDownloadButton > button {
  border-radius: 10px; font-weight: 600; min-height: 2.7rem;
}
div[data-testid="stForm"] {
  background: white; border: 1px solid #e4ebf4; border-radius: 18px;
  padding: 1.2rem 1.3rem;
}
</style>
<div class="hero">
  <h1>🩺 LambungCare</h1>
  <p>Sistem Pendukung Diagnosis Penyakit Lambung · IVIFS, IIFWAA, dan Jarak Hamming</p>
</div>
""", unsafe_allow_html=True)

st.warning(
    "Aplikasi ini merupakan prototipe penelitian, bukan pengganti pemeriksaan dokter. "
    "Hasil sistem bukan diagnosis pasti dan tidak boleh menjadi satu-satunya dasar tindakan medis."
)

# --------------------------- Fungsi metode ---------------------------
def format_interval(a, b):
    return f"[{a:.4f}; {b:.4f}]"

def bangun_acuan(lebar=LEBAR_INTERVAL, delta=DELTA):
    ivifs_awal, ivifs_normalisasi = {}, {}
    for penyakit, skor_gejala in DATA_ACUAN.items():
        ivifs_awal[penyakit], ivifs_normalisasi[penyakit] = {}, {}
        for gejala, skor in zip(GEJALA, skor_gejala):
            x = skor / 10.0
            nilai = [max(0.0, x - lebar), min(1.0, x + lebar)] if lebar > 0 else [x]
            anggota = [v for v in nilai if v > 0.5]
            takanggota = [v for v in nilai if v < 0.5]
            s_l, s_u = (min(anggota), max(anggota)) if anggota else (0.5, 0.5)
            n_l, n_u = (min(takanggota), max(takanggota)) if takanggota else (0.5, 0.5)
            awal = (2*s_l-1, 2*s_u-1, 1-2*n_u, 1-2*n_l)
            ivifs_awal[penyakit][gejala] = awal
            ivifs_normalisasi[penyakit][gejala] = tuple(v / delta for v in awal)
    return ivifs_awal, ivifs_normalisasi

def skor_pasien_ke_ivifs(skor, penyebut=DELTA):
    if not 0 <= int(skor) <= 10:
        raise ValueError(f"Skor harus 0–10, didapat {skor}")
    x = int(skor) / 10.0
    if x > 0.5:
        s, n = x, 0.5
    elif x < 0.5:
        s, n = 0.5, x
    else:
        s, n = 0.5, 0.5
    return ((2*s-1)/penyebut, (2*s-1)/penyebut,
            (1-2*n)/penyebut, (1-2*n)/penyebut)

def agregasi_iifwaa(daftar_ivifs, daftar_bobot=BOBOT):
    total = sum(daftar_bobot)
    if total <= 0:
        raise ValueError("Jumlah bobot harus lebih besar dari 0.")
    w = [b/total for b in daftar_bobot]
    mu_l = 1 - math.prod((1-x[0])**wi for x, wi in zip(daftar_ivifs, w))
    mu_u = 1 - math.prod((1-x[1])**wi for x, wi in zip(daftar_ivifs, w))
    nu_l = math.prod(x[2]**wi for x, wi in zip(daftar_ivifs, w))
    nu_u = math.prod(x[3]**wi for x, wi in zip(daftar_ivifs, w))
    return mu_l, mu_u, nu_l, nu_u

def hesitasi(ivifs):
    mu_l, mu_u, nu_l, nu_u = ivifs
    return 1-mu_u-nu_u, 1-mu_l-nu_l

def jarak_hamming(a, b):
    pi_a, pi_b = hesitasi(a), hesitasi(b)
    return (abs(a[0]-b[0]) + abs(a[1]-b[1]) +
            abs(a[2]-b[2]) + abs(a[3]-b[3]) +
            abs(pi_a[0]-pi_b[0]) + abs(pi_a[1]-pi_b[1])) / 6.0

def hitung_diagnosis(skor_pasien):
    _, acuan = bangun_acuan()
    ivifs_pasien = [skor_pasien_ke_ivifs(s) for s in skor_pasien]
    baris_ivifs = []
    for gejala, skor, iv in zip(GEJALA, skor_pasien, ivifs_pasien):
        pi = hesitasi(iv)
        baris_ivifs.append({
            "Gejala": gejala, "Skor pasien": int(skor),
            "Keanggotaan": format_interval(iv[0], iv[1]),
            "Ketakanggotaan": format_interval(iv[2], iv[3]),
            "Hesitasi": format_interval(pi[0], pi[1]),
        })
    agregat_pasien = agregasi_iifwaa(ivifs_pasien)
    pi = hesitasi(agregat_pasien)
    tabel_agregat_pasien = pd.DataFrame([{
        "Keanggotaan": format_interval(agregat_pasien[0], agregat_pasien[1]),
        "Ketakanggotaan": format_interval(agregat_pasien[2], agregat_pasien[3]),
        "Hesitasi": format_interval(pi[0], pi[1]),
    }])
    agregat_acuan_rows, jarak_per_penyakit = [], {}
    for p in PENYAKIT:
        daftar = [acuan[p][g] for g in GEJALA]
        ag = agregasi_iifwaa(daftar)
        hp = hesitasi(ag)
        agregat_acuan_rows.append({
            "Penyakit": p, "Keanggotaan": format_interval(ag[0], ag[1]),
            "Ketakanggotaan": format_interval(ag[2], ag[3]),
            "Hesitasi": format_interval(hp[0], hp[1]),
        })
        jarak_per_penyakit[p] = float(np.mean([
            jarak_hamming(ivifs_pasien[i], acuan[p][g])
            for i, g in enumerate(GEJALA)
        ]))
    urutan = sorted(jarak_per_penyakit.items(), key=lambda kv: kv[1])
    tabel_jarak = pd.DataFrame([{
        "Peringkat": i+1, "Penyakit": p, "Jarak Hamming rata-rata": d,
        "Keterangan": "Jarak terkecil" if i == 0 else ""
    } for i, (p, d) in enumerate(urutan)])
    return {
        "diagnosis": urutan[0][0],
        "tabel_ivifs": pd.DataFrame(baris_ivifs),
        "tabel_jarak": tabel_jarak,
        "agregat_pasien": tabel_agregat_pasien,
        "agregat_acuan": pd.DataFrame(agregat_acuan_rows),
    }

def buat_excel(riwayat):
    output = io.BytesIO()
    data_pasien = []
    ivifs_rows, jarak_rows, agregat_pasien_rows = [], [], []
    for item in riwayat:
        base = {
            "Waktu pemeriksaan": item["waktu"], "Kode pasien": item["kode"],
            "Usia": item["usia"], "Jenis kelamin": item["jk"],
            **{g: int(s) for g, s in zip(GEJALA, item["skor"])},
            "Hasil sistem": item["hasil"]["diagnosis"],
        }
        data_pasien.append(base)
        for _, r in item["hasil"]["tabel_ivifs"].iterrows():
            ivifs_rows.append({"Kode pasien": item["kode"], **r.to_dict()})
        for _, r in item["hasil"]["tabel_jarak"].iterrows():
            jarak_rows.append({"Kode pasien": item["kode"], **r.to_dict()})
        for _, r in item["hasil"]["agregat_pasien"].iterrows():
            agregat_pasien_rows.append({"Kode pasien": item["kode"], **r.to_dict()})
    data_acuan = pd.DataFrame([
        {"Penyakit": p, "Gejala": g, "Skor acuan Dokter 3": s}
        for p in PENYAKIT for g, s in zip(GEJALA, DATA_ACUAN[p])
    ])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(data_pasien).to_excel(writer, sheet_name="Rekap Pasien", index=False)
        pd.DataFrame(ivifs_rows).to_excel(writer, sheet_name="Konversi IVIFS", index=False)
        pd.DataFrame(jarak_rows).to_excel(writer, sheet_name="Jarak Hamming", index=False)
        pd.DataFrame(agregat_pasien_rows).to_excel(writer, sheet_name="Agregat Pasien", index=False)
        data_acuan.to_excel(writer, sheet_name="Data Acuan", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len+2, 12), 42)
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
    output.seek(0)
    return output.getvalue()

# --------------------------- Penyimpanan JSON ---------------------------
# File ini membantu menyimpan riwayat antar-rerun aplikasi.
# Pada hosting cloud, file lokal dapat hilang saat server/redeploy direset.
FILE_RIWAYAT_JSON = Path(__file__).with_name("riwayat_pasien.json")

def simpan_riwayat_json(riwayat):
    """Simpan data input mentah; hasil perhitungan dibuat ulang saat dimuat."""
    data_simpan = []
    for item in riwayat:
        data_simpan.append({
            "waktu": item["waktu"],
            "kode": item["kode"],
            "usia": item["usia"],
            "jk": item["jk"],
            "skor": [int(s) for s in item["skor"]],
        })
    FILE_RIWAYAT_JSON.write_text(
        json.dumps(data_simpan, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

def muat_riwayat_json():
    """Muat riwayat JSON dan bangun ulang tabel hasil untuk setiap pasien."""
    if not FILE_RIWAYAT_JSON.exists():
        return []
    try:
        data = json.loads(FILE_RIWAYAT_JSON.read_text(encoding="utf-8"))
        riwayat = []
        for item in data:
            skor = [int(s) for s in item["skor"]]
            if len(skor) != len(GEJALA):
                continue
            riwayat.append({
                "waktu": item.get("waktu", ""),
                "kode": str(item.get("kode", "")),
                "usia": item.get("usia", ""),
                "jk": item.get("jk", "P"),
                "skor": skor,
                "hasil": hitung_diagnosis(skor),
            })
        return riwayat
    except (OSError, json.JSONDecodeError, KeyError, TypeError, ValueError) as e:
        st.warning(f"Riwayat JSON belum bisa dibaca: {e}")
        return []

# --------------------------- Session state ---------------------------
if "riwayat_pasien" not in st.session_state:
    st.session_state.riwayat_pasien = muat_riwayat_json()
if "nomor_form" not in st.session_state:
    st.session_state.nomor_form = 1
if "hasil_terakhir" not in st.session_state:
    st.session_state.hasil_terakhir = None

with st.sidebar:
    st.markdown("## 📋 Panel pemeriksaan")
    st.metric("Pasien sudah diproses", len(st.session_state.riwayat_pasien))
    st.caption("Riwayat dimuat dari file JSON jika file tersebut masih tersedia di server.")
    st.divider()
    st.markdown("### Konfigurasi metode")
    st.write("**Data acuan:** Dokter 3")
    st.write(f"**Lebar interval:** ±{LEBAR_INTERVAL:.2f}")
    st.write(f"**Normalisasi Δ:** {DELTA:.2f}")
    st.write("**Bobot gejala:** sama (1/7)")
    if st.button("🗑️ Hapus seluruh riwayat", use_container_width=True):
        st.session_state.riwayat_pasien = []
        simpan_riwayat_json([])
        st.session_state.hasil_terakhir = None
        st.session_state.nomor_form += 1
        st.rerun()

tab_input, tab_riwayat, tab_acuan = st.tabs(["➕ Pemeriksaan pasien", "📚 Riwayat & unduh Excel", "🧪 Data acuan"])

with tab_input:
    st.markdown("### Pemeriksaan pasien")
    st.markdown(
        f'<p class="small-muted">Formulir #{st.session_state.nomor_form} · '
        'Isi identitas dengan kode pasien, lalu beri skor untuk setiap gejala.</p>',
        unsafe_allow_html=True,
    )
    with st.form(f"form_pasien_{st.session_state.nomor_form}"):
        st.markdown("#### 1. Identitas pasien")
        c1, c2, c3 = st.columns(3)
        with c1:
            kode = st.text_input("Kode pasien *", value=f"P{len(st.session_state.riwayat_pasien)+1:03d}",
                                 max_chars=30, key=f"kode_{st.session_state.nomor_form}")
        with c2:
            usia_teks = st.text_input(
                "Usia (tahun)",
                value="",
                placeholder="Masukkan usia",
                key=f"usia_{st.session_state.nomor_form}",
                help="Boleh dikosongkan jika usia tidak tersedia."
            )
        with c3:
            jk = st.selectbox("Jenis kelamin", ["P", "L"],
                              format_func=lambda x: "Perempuan (P)" if x == "P" else "Laki-laki (L)",
                              key=f"jk_{st.session_state.nomor_form}")
        st.markdown("#### 2. Penilaian gejala")
        st.caption("Skala 0–10: 0 berarti gejala tidak dirasakan, sedangkan 10 berarti gejala sangat berat.")
        skor_pasien = []
        cols = st.columns(2)
        for i, gejala in enumerate(GEJALA):
            with cols[i % 2]:
                skor_pasien.append(st.slider(
                    gejala, min_value=0, max_value=10, value=0, step=1,
                    key=f"skor_{st.session_state.nomor_form}_{i}",
                    help="Pilih 0 jika gejala tidak dirasakan; skor 0–10 sesuai protokol penelitian."
                ))
        proses = st.form_submit_button("🔎 Proses pemeriksaan pasien", type="primary", use_container_width=True)

    if proses:
        kode_bersih = kode.strip()
        if not kode_bersih:
            st.error("Kode pasien wajib diisi.")
        elif any(x["kode"].casefold() == kode_bersih.casefold() for x in st.session_state.riwayat_pasien):
            st.error("Kode pasien ini sudah ada di riwayat. Gunakan kode berbeda agar data tidak tertukar.")
        else:
            try:
                hasil = hitung_diagnosis(skor_pasien)
                usia_bersih = usia_teks.strip()
                if usia_bersih:
                    if not usia_bersih.isdigit() or not 0 <= int(usia_bersih) <= 120:
                        st.error("Usia harus berupa angka 0–120, atau dikosongkan jika tidak diketahui.")
                        st.stop()
                    usia_simpan = int(usia_bersih)
                else:
                    usia_simpan = ""

                item = {
                    "waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "kode": kode_bersih, "usia": usia_simpan, "jk": jk,
                    "skor": list(skor_pasien), "hasil": hasil,
                }
                st.session_state.riwayat_pasien.append(item)
                simpan_riwayat_json(st.session_state.riwayat_pasien)
                st.session_state.hasil_terakhir = item
                st.success(f"Data pasien {kode_bersih} berhasil diproses dan ditambahkan ke riwayat.")
            except Exception as e:
                st.error(f"Perhitungan gagal: {e}")

    # Hasil pemeriksaan terbaru
    terbaru = st.session_state.hasil_terakhir
    if terbaru:
        st.divider()
        st.markdown("### Hasil pemeriksaan terakhir")
        m1, m2, m3 = st.columns(3)
        m1.metric("Kode pasien", terbaru["kode"])
        m2.metric("Prediksi terdekat model", terbaru["hasil"]["diagnosis"])
        m3.metric("Gejala dinilai", f"{len(terbaru['skor'])}/{len(GEJALA)}")
        st.caption("Prediksi merupakan penyakit dengan jarak Hamming rata-rata terkecil terhadap data acuan yang tertanam.")
        st.dataframe(
            terbaru["hasil"]["tabel_jarak"].style.format({"Jarak Hamming rata-rata": "{:.6f}"}),
            use_container_width=True, hide_index=True
        )
        with st.expander("Lihat detail konversi IVIFS pasien"):
            st.dataframe(terbaru["hasil"]["tabel_ivifs"], use_container_width=True, hide_index=True)
        with st.expander("Lihat agregat IIFWAA pasien"):
            st.dataframe(terbaru["hasil"]["agregat_pasien"], use_container_width=True, hide_index=True)
        b1, b2 = st.columns(2)
        with b1:
            if st.button("➡️ Lanjut ke pasien berikutnya", type="primary", use_container_width=True):
                st.session_state.nomor_form += 1
                st.session_state.hasil_terakhir = None
                st.rerun()
        with b2:
            st.download_button(
                "⬇️ Unduh Excel semua pasien",
                data=buat_excel(st.session_state.riwayat_pasien),
                file_name=f"Rekap_Diagnosis_Lambung_{datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.info("Isi formulir lalu tekan **Proses pemeriksaan pasien**. Setelah hasil muncul, pilih **Lanjut ke pasien berikutnya** untuk memulai pasien baru.")

with tab_riwayat:
    st.markdown("### Riwayat pemeriksaan")
    if st.session_state.riwayat_pasien:
        rows = []
        for x in st.session_state.riwayat_pasien:
            rows.append({
                "Waktu pemeriksaan": x["waktu"], "Kode pasien": x["kode"],
                "Usia": x["usia"] if x["usia"] != "" else "—",
                "Jenis kelamin": "Perempuan" if x["jk"] == "P" else "Laki-laki",
                "Hasil sistem": x["hasil"]["diagnosis"],
                **{g: s for g, s in zip(GEJALA, x["skor"])},
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        st.download_button(
            "📥 Download rekap lengkap ke Excel (.xlsx)",
            data=buat_excel(st.session_state.riwayat_pasien),
            file_name=f"Rekap_Diagnosis_Lambung_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
        st.caption("File Excel berisi rekap pasien, konversi IVIFS tiap pasien, jarak Hamming, agregat pasien, dan data acuan.")
    else:
        st.info("Belum ada pasien yang diproses. Data akan muncul di sini setelah pemeriksaan pertama selesai.")

with tab_acuan:
    st.markdown("### Data acuan penyakit")
    st.caption("Skor acuan Dokter 3 yang tertanam di kode sumber. Pastikan kesesuaiannya dengan data penelitian.")
    st.dataframe(pd.DataFrame(DATA_ACUAN, index=GEJALA).T.rename_axis("Penyakit"),
                 use_container_width=True)
    st.markdown("### Catatan penting")
    st.write(
        "Data riwayat disimpan di memori sesi aplikasi dan dapat diunduh ke Excel. "
        "Jika server/aplikasi dimulai ulang atau sesi berakhir, riwayat sementara dapat hilang. "
        "Simpan file Excel secara berkala. Untuk penggunaan banyak pengguna atau penyimpanan permanen, "
        "perlu ditambahkan basis data."
    )
    st.info(
        "Rumus konversi IVIFS, definisi jarak Hamming, konstanta Δ=1,6, dan data acuan perlu "
        "divalidasi terhadap jurnal sumber serta dokter sebelum aplikasi digunakan untuk penelitian/layanan nyata. "
        "Hindari memasukkan nama lengkap atau identitas langsung pasien."
    )
