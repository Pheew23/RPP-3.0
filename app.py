"""
Generator Modul Ajar berbasis AI
---------------------------------
Streamlit app yang memakai model AI dari NVIDIA build (build.nvidia.com / NIM)
untuk menyusun isi Modul Ajar (Kurikulum Berbasis Cinta - Deep Learning)
lengkap dengan identifikasi peserta didik, desain pembelajaran, kegiatan per
pertemuan, asesmen, materi ajar, dan tindak lanjut - lalu diekspor ke .docx.

Jalankan:
    pip install -r requirements.txt
    streamlit run app.py

API key disimpan di .streamlit/secrets.toml (lihat secrets.toml.example),
BUKAN ditulis langsung di kode.
"""

import io
import json
import re

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Palet warna diambil langsung dari file contoh Modul Ajar (fill hex asli)
COLOR_TITLE = "1F4E79"       # kotak judul modul (biru tua)
COLOR_IDENTITY_HEAD = "2E75B6"   # header tabel identitas / pengalaman belajar
COLOR_LABEL = "DEEAF1"       # kolom label (biru muda)
COLOR_VALUE = "FFFFFF"       # kolom isi (putih)
COLOR_SECTION_A = "1F4E79"   # A. Identifikasi Peserta Didik
COLOR_SECTION_B = "375623"   # C. Desain Pembelajaran (hijau tua)
COLOR_MEETING = "843C0C"     # Pengalaman Belajar per pertemuan (coklat/rust)
COLOR_LAMPIRAN_I = "C00000"  # Lampiran I - Asesmen (merah)
COLOR_LAMPIRAN_II = "375623"  # Lampiran II - Materi Ajar (hijau tua)
COLOR_LAMPIRAN_III = "006060"  # Lampiran III - LKPD (teal)
COLOR_LAMPIRAN_V = "2E75B6"  # Lampiran V - Tindak Lanjut (biru)
COLOR_REMEDIAL = "C00000"
COLOR_PENGAYAAN = "375623"
COLOR_REFLEKSI = "2E75B6"

MODEL_NAME = "nvidia/nemotron-3-ultra-550b-a55b"
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"

FASE_MAP = {1: "A", 2: "A", 3: "B", 4: "B", 5: "C", 6: "C"}

st.set_page_config(page_title="Generator Modul Ajar AI", page_icon="📘", layout="wide")

# ---------------------------------------------------------------------------
# Klien NVIDIA (OpenAI-compatible endpoint)
# ---------------------------------------------------------------------------
@st.cache_resource
def get_client():
    api_key = st.secrets.get("NVIDIA_API_KEY")
    if not api_key:
        st.error(
            "NVIDIA_API_KEY belum ada di st.secrets. "
            "Tambahkan di .streamlit/secrets.toml."
        )
        st.stop()
    return OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)


SCHEMA_HINT = """{
  "capaian_pembelajaran": "string",
  "tujuan_pembelajaran": ["string", "string", "string", "string"],
  "lintas_disiplin": ["string", "string", "string"],
  "identifikasi": {
    "pengetahuan_awal": ["string", "string", "string", "string"],
    "minat_belajar": ["string", "string", "string", "string"],
    "latar_belakang": "string paragraf singkat",
    "kebutuhan_belajar": ["string", "string", "string", "string"],
    "dimensi_profil_kelulusan": ["string", "string", "string", "string"],
    "topik_panca_cinta": ["string", "string", "string", "string", "string"]
  },
  "asesmen_awal": ["string", "string", "string"],
  "materi_ajar": "string 2-4 paragraf",
  "pertemuan": [
    {
      "nomor": 1,
      "tujuan": "string",
      "pembuka": ["string", "string"],
      "inti": ["string", "string", "string", "string"],
      "penutup": ["string", "string"],
      "lkpd_judul": "string",
      "lkpd_instruksi": ["string", "string", "string"]
    }
  ],
  "asesmen_sumatif": ["string", "string", "string", "string", "string"],
  "remedial": "string",
  "pengayaan": "string",
  "refleksi": ["string", "string", "string"]
}"""


def build_prompt(form: dict) -> str:
    fase = FASE_MAP[form["kelas"]]
    n = form["jumlah_pertemuan"]
    return f"""Kamu adalah pakar kurikulum sekolah dasar/madrasah Indonesia.
Buat ISI Modul Ajar (Kurikulum Berbasis Cinta - Pendekatan Deep Learning) dalam
Bahasa Indonesia untuk:
- Mata Pelajaran: {form['mapel']}
- Kelas / Fase: {form['kelas']} / Fase {fase}
- Semester: {form['semester']}
- Bab / Topik: {form['bab']}
- Jumlah Pertemuan: {n}
- Model Pembelajaran: {form['model']}

Buat array "pertemuan" berisi TEPAT {n} objek (nomor 1 sampai {n}), masing-masing
dengan kegiatan pembuka/inti/penutup dan LKPD yang berbeda dan berkembang secara
logis dari pertemuan sebelumnya, semuanya relevan dengan bab/topik di atas.

Balas HANYA dengan JSON valid sesuai skema berikut. Jangan tambahkan teks
pembuka, penjelasan, atau markdown code fence:
{SCHEMA_HINT}"""


def call_ai(form: dict) -> dict:
    client = get_client()
    prompt = build_prompt(form)
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.6,
        max_tokens=6000,
    )
    text = response.choices[0].message.content.strip()
    
    # PERBAIKAN: Menggunakan regex untuk mengekstrak hanya format JSON
    # Ini mencegah error jika AI memberikan output "Berikut adalah hasilnya: { ... }"
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group(0)
        
    return json.loads(text)


# ---------------------------------------------------------------------------
# Ekspor .docx - meniru skema warna file contoh
# ---------------------------------------------------------------------------
def set_cell_background(cell, hex_color: str):
    """Beri warna latar (shading) pada sebuah cell tabel."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def style_cell_text(cell, text, bold=False, color_hex=None, size=11, center=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    if color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)


def banner(doc, text, hex_color, size=12):
    """Baris pita warna penuh selebar halaman, teks putih tebal - dipakai
    untuk judul section."""
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    table.columns[0].width = Cm(17)
    cell = table.rows[0].cells[0]
    set_cell_background(cell, hex_color)
    style_cell_text(cell, text, bold=True, color_hex="FFFFFF", size=size)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


def field_table(doc):
    """Tabel 2 kolom (label | isi) yang lebar labelnya konsisten."""
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(4.5)
    table.columns[1].width = Cm(12.5)
    return table


def add_field_row(table, label, content_items):
    """Satu baris: kolom kiri label (biru muda), kolom kanan isi (putih)."""
    row = table.add_row()
    label_cell, value_cell = row.cells[0], row.cells[1]
    label_cell.width = Cm(4.5)
    value_cell.width = Cm(12.5)
    set_cell_background(label_cell, COLOR_LABEL)
    set_cell_background(value_cell, COLOR_VALUE)
    style_cell_text(label_cell, label, bold=True, size=10.5)

    value_cell.text = ""
    if isinstance(content_items, (list, tuple)):
        first = True
        for item in content_items:
            p = value_cell.paragraphs[0] if first else value_cell.add_paragraph()
            p.style = doc_bullet_style(value_cell)
            run = p.add_run(f"\u2022 {item}")
            run.font.size = Pt(10.5)
            first = False
    else:
        p = value_cell.paragraphs[0]
        run = p.add_run(str(content_items))
        run.font.size = Pt(10.5)
    return row


def doc_bullet_style(cell):
    return cell.paragraphs[0].style


def build_docx(form: dict, data: dict) -> bytes:
    fase = FASE_MAP[form["kelas"]]
    doc = Document()
    doc.sections[0].left_margin = Cm(2)
    doc.sections[0].right_margin = Cm(2)
    doc.sections[0].page_width = Cm(21)

    # --- Kotak judul biru tua ---
    title_table = doc.add_table(rows=1, cols=1)
    cell = title_table.rows[0].cells[0]
    set_cell_background(cell, COLOR_TITLE)
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run("MODUL AJAR")
    r1.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor.from_string("FFFFFF")
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("KURIKULUM BERBASIS CINTA \u2013 PENDEKATAN DEEP LEARNING")
    r2.bold = True
    r2.font.size = Pt(11)
    r2.font.color.rgb = RGBColor.from_string("FFFFFF")
    doc.add_paragraph()

    # --- Header identitas ---
    banner(doc, "IDENTITAS MODUL AJAR", COLOR_IDENTITY_HEAD)
    identity = field_table(doc)
    for label, value in [
        ("Mata Pelajaran", form["mapel"]),
        ("Kelas / Fase", f"Kelas {form['kelas']} / Fase {fase}"),
        ("Semester", form["semester"]),
        ("Alokasi Waktu", f"{len(data['pertemuan'])} Pertemuan x {form['alokasi']}"),
        ("Bab / Topik", form["bab"]),
        ("Model Pembelajaran", form["model"]),
        ("Metode Pembelajaran", form["metode"]),
        ("Penyusun", form["penyusun"]),
        ("Sekolah", form["sekolah"]),
        ("Tahun Pelajaran", form["tahun_pelajaran"]),
    ]:
        add_field_row(identity, label, value)
    doc.add_paragraph()

    # --- A. Identifikasi Peserta Didik ---
    banner(doc, "A. IDENTIFIKASI PESERTA DIDIK", COLOR_SECTION_A)
    ident = field_table(doc)
    add_field_row(ident, "Pengetahuan Awal", data["identifikasi"]["pengetahuan_awal"])
    add_field_row(ident, "Minat Belajar", data["identifikasi"]["minat_belajar"])
    add_field_row(ident, "Latar Belakang", data["identifikasi"]["latar_belakang"])
    add_field_row(ident, "Kebutuhan Belajar", data["identifikasi"]["kebutuhan_belajar"])
    add_field_row(ident, "Dimensi Profil Kelulusan", data["identifikasi"]["dimensi_profil_kelulusan"])
    add_field_row(ident, "Topik Panca Cinta", data["identifikasi"]["topik_panca_cinta"])
    doc.add_paragraph()

    # --- B/C. Desain Pembelajaran ---
    banner(doc, "C. DESAIN PEMBELAJARAN", COLOR_SECTION_B)
    desain = field_table(doc)
    add_field_row(desain, "Capaian Pembelajaran (CP)", data["capaian_pembelajaran"])
    add_field_row(
        desain, "Tujuan Pembelajaran (TP)",
        [f"TP {i}: {tp}" for i, tp in enumerate(data["tujuan_pembelajaran"], 1)],
    )
    add_field_row(desain, "Lintas Disiplin Ilmu", data["lintas_disiplin"])
    doc.add_paragraph()

    # --- Pengalaman Belajar per pertemuan ---
    for p in data["pertemuan"]:
        banner(doc, f"PENGALAMAN BELAJAR \u2013 PERTEMUAN {p['nomor']}", COLOR_MEETING, size=11)
        meeting = field_table(doc)
        add_field_row(meeting, "Tujuan Pertemuan", p["tujuan"])
        add_field_row(meeting, "Kegiatan Pembuka", p["pembuka"])
        add_field_row(meeting, "Kegiatan Inti", p["inti"])
        add_field_row(meeting, "Kegiatan Penutup", p["penutup"])
        doc.add_paragraph()

    # --- Lampiran I: Asesmen ---
    banner(doc, "LAMPIRAN I \u2013 ASESMEN", COLOR_LAMPIRAN_I)
    banner(doc, "A. ASESMEN AWAL (LISAN)", COLOR_LAMPIRAN_I, size=10.5)
    asesmen_awal = field_table(doc)
    add_field_row(asesmen_awal, "Pertanyaan Lisan", data["asesmen_awal"])
    banner(doc, "D. ASESMEN SUMATIF \u2013 5 SOAL URAIAN HOTS", COLOR_MEETING, size=10.5)
    sumatif = field_table(doc)
    add_field_row(
        sumatif, "Soal",
        [f"{i}. {s}" for i, s in enumerate(data["asesmen_sumatif"], 1)],
    )
    doc.add_paragraph()

    # --- Lampiran II: Materi Ajar ---
    banner(doc, "LAMPIRAN II \u2013 MATERI AJAR", COLOR_LAMPIRAN_II)
    doc.add_paragraph(data["materi_ajar"])
    doc.add_paragraph()

    # --- Lampiran III: LKPD ---
    banner(doc, "LAMPIRAN III \u2013 LKPD (LEMBAR KERJA PESERTA DIDIK)", COLOR_LAMPIRAN_III)
    for p in data["pertemuan"]:
        banner(doc, f"LKPD PERTEMUAN {p['nomor']} \u2013 {p['lkpd_judul']}", COLOR_LAMPIRAN_III, size=10.5)
        lkpd = field_table(doc)
        add_field_row(lkpd, "Instruksi", p["lkpd_instruksi"])
        doc.add_paragraph()

    # --- Lampiran V: Tindak Lanjut dan Refleksi ---
    banner(doc, "LAMPIRAN V \u2013 TINDAK LANJUT DAN REFLEKSI", COLOR_LAMPIRAN_V)
    banner(doc, "A. PROGRAM REMEDIAL", COLOR_REMEDIAL, size=10.5)
    doc.add_paragraph(data["remedial"])
    banner(doc, "B. PROGRAM PENGAYAAN", COLOR_PENGAYAAN, size=10.5)
    doc.add_paragraph(data["pengayaan"])
    banner(doc, "C. INSTRUMEN REFLEKSI PEMBELAJARAN", COLOR_REFLEKSI, size=10.5)
    refleksi = field_table(doc)
    add_field_row(refleksi, "Pertanyaan Refleksi", data["refleksi"])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📘 Generator Modul Ajar Guru (AI - NVIDIA Build)")
st.caption(f"Model: `{MODEL_NAME}` via build.nvidia.com")

with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Bahasa Indonesia")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Anak-Anak yang Mengubah Dunia")
        kelas = st.selectbox("Kelas", [1, 2, 3, 4, 5, 6], format_func=lambda k: f"Kelas {k} (Fase {FASE_MAP[k]})")
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan", min_value=1, max_value=8, value=4)
    with col2:
        alokasi = st.text_input("Alokasi Waktu per Pertemuan", value="4 JP x 35 menit")
        model_pembelajaran = st.text_input("Model Pembelajaran", value="PBL (Problem Based Learning)")
        metode = st.text_input("Metode Pembelajaran", value="Ceramah Interaktif, Diskusi Kelompok, Tanya Jawab")
        penyusun = st.text_input("Penyusun (Nama Guru)")
        sekolah = st.text_input("Sekolah")
        tahun_pelajaran = st.text_input("Tahun Pelajaran", value="2026/2027")

    submitted = st.form_submit_button("✨ Buat Modul Ajar dengan AI", use_container_width=True)

if submitted:
    if not (mapel and bab and penyusun and sekolah):
        st.warning("Lengkapi minimal Mata Pelajaran, Bab/Topik, Penyusun, dan Sekolah.")
    else:
        form = dict(
            mapel=mapel, bab=bab, kelas=kelas, semester=semester,
            jumlah_pertemuan=int(jumlah_pertemuan), alokasi=alokasi,
            model=model_pembelajaran, metode=metode, penyusun=penyusun,
            sekolah=sekolah, tahun_pelajaran=tahun_pelajaran,
        )
        with st.spinner("AI sedang menyusun modul ajar..."):
            try:
                data = call_ai(form)
                st.session_state["form"] = form
                st.session_state["data"] = data
            except json.JSONDecodeError:
                st.error("AI mengembalikan format JSON yang tidak valid. Coba klik tombol sekali lagi.")
            except Exception as e:
                st.error(f"Gagal memanggil model NVIDIA: {e}")

if "data" in st.session_state:
    form = st.session_state["form"]
    data = st.session_state["data"]
    fase = FASE_MAP[form["kelas"]]

    st.divider()
    st.subheader(f"Pratinjau: {form['mapel']} \u2013 {form['bab']}")

    docx_bytes = build_docx(form, data)
    st.download_button(
        "⬇️ Unduh sebagai Word (.docx)",
        data=docx_bytes,
        file_name=f"Modul_Ajar_{form['mapel'].replace(' ', '_')}_Kelas{form['kelas']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    with st.expander("A. Identifikasi Peserta Didik", expanded=True):
        st.markdown("**Pengetahuan Awal**")
        for s in data["identifikasi"]["pengetahuan_awal"]:
            st.markdown(f"- {s}")
        st.markdown("**Minat Belajar**")
        for s in data["identifikasi"]["minat_belajar"]:
            st.markdown(f"- {s}")
        st.markdown("**Latar Belakang**")
        st.write(data["identifikasi"]["latar_belakang"])
        st.markdown("**Kebutuhan Belajar**")
        for s in data["identifikasi"]["kebutuhan_belajar"]:
            st.markdown(f"- {s}")
        st.markdown("**Dimensi Profil Kelulusan**")
        for s in data["identifikasi"]["dimensi_profil_kelulusan"]:
            st.markdown(f"- {s}")
        st.markdown("**Topik Panca Cinta**")
        for s in data["identifikasi"]["topik_panca_cinta"]:
            st.markdown(f"- {s}")

    with st.expander("B. Desain Pembelajaran"):
        st.markdown("**Capaian Pembelajaran**")
        st.write(data["capaian_pembelajaran"])
        st.markdown("**Tujuan Pembelajaran**")
        for i, tp in enumerate(data["tujuan_pembelajaran"], 1):
            st.markdown(f"- TP {i}: {tp}")
        st.markdown("**Lintas Disiplin Ilmu**")
        for s in data["lintas_disiplin"]:
            st.markdown(f"- {s}")

    with st.expander("C. Pengalaman Belajar per Pertemuan"):
        for p in data["pertemuan"]:
            st.markdown(f"#### Pertemuan {p['nomor']} \u2013 {p['tujuan']}")
            st.markdown("**Pembuka**")
            for s in p["pembuka"]:
                st.markdown(f"- {s}")
            st.markdown("**Inti**")
            for s in p["inti"]:
                st.markdown(f"- {s}")
            st.markdown("**Penutup**")
            for s in p["penutup"]:
                st.markdown(f"- {s}")
            st.markdown(f"**LKPD: {p['lkpd_judul']}**")
            for s in p["lkpd_instruksi"]:
                st.markdown(f"- {s}")
            st.markdown("---")

    with st.expander("Lampiran I \u2013 Asesmen"):
        st.markdown("**Asesmen Awal (Lisan)**")
        for s in data["asesmen_awal"]:
            st.markdown(f"- {s}")
        st.markdown("**Asesmen Sumatif \u2013 5 Soal Uraian HOTS**")
        for i, s in enumerate(data["asesmen_sumatif"], 1):
            st.markdown(f"{i}. {s}")

    with st.expander("Lampiran II \u2013 Materi Ajar"):
        st.write(data["materi_ajar"])

    with st.expander("Lampiran V \u2013 Tindak Lanjut dan Refleksi"):
        st.markdown("**Program Remedial**")
        st.write(data["remedial"])
        st.markdown("**Program Pengayaan**")
        st.write(data["pengayaan"])
        st.markdown("**Instrumen Refleksi**")
        for s in data["refleksi"]:
            st.markdown(f"- {s}")
