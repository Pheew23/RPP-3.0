"""
Generator Modul Ajar berbasis AI - KBC, KMA 1503/2025 & Auto Download
--------------------------------------------------------------------------------
Mendukung RA/TK hingga SMA/MA, referensi KMA 1503 Tahun 2025, Prompt Chaining, 
Desain Tabel Berwarna, Tanda Tangan Pengesahan, dan Auto Pop-up Download.
"""

import io
import json
import re
import datetime
import base64

import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==============================================================================
# PALET WARNA
# ==============================================================================
COLOR_TITLE = "1F4E79"       
COLOR_IDENTITY_HEAD = "2E75B6"   
COLOR_LABEL = "DEEAF1"       
COLOR_VALUE = "FFFFFF"       
COLOR_SECTION_A = "1F4E79"   
COLOR_SECTION_B = "375623"   
COLOR_MEETING = "843C0C"     
COLOR_LAMPIRAN_I = "C00000"  
COLOR_LAMPIRAN_II = "375623"  
COLOR_LAMPIRAN_III = "006060"  
COLOR_LAMPIRAN_V = "2E75B6"  
COLOR_REMEDIAL = "C00000"
COLOR_PENGAYAAN = "375623"
COLOR_REFLEKSI = "2E75B6"

MODEL_NAME = "nvidia/nemotron-3-ultra-550b-a55b"
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"

# Peta Jenjang & Fase Lengkap
JENJANG_FASE = {
    "RA/TK (Fase Fondasi)": "Fondasi",
    "Kelas 1 SD/MI (Fase A)": "A",
    "Kelas 2 SD/MI (Fase A)": "A",
    "Kelas 3 SD/MI (Fase B)": "B",
    "Kelas 4 SD/MI (Fase B)": "B",
    "Kelas 5 SD/MI (Fase C)": "C",
    "Kelas 6 SD/MI (Fase C)": "C",
    "Kelas 7 SMP/MTs (Fase D)": "D",
    "Kelas 8 SMP/MTs (Fase D)": "D",
    "Kelas 9 SMP/MTs (Fase D)": "D",
    "Kelas 10 SMA/MA/SMK (Fase E)": "E",
    "Kelas 11 SMA/MA/SMK (Fase F)": "F",
    "Kelas 12 SMA/MA/SMK (Fase F)": "F",
}

st.set_page_config(page_title="MIFSAL RPP V3", page_icon="📘", layout="wide")

@st.cache_resource
def get_client():
    api_key = st.secrets.get("NVIDIA_API_KEY")
    if not api_key:
        st.error("NVIDIA_API_KEY belum ada di st.secrets.")
        st.stop()
    return OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)

def call_ai(prompt: str, temperature=0.2) -> dict:
    client = get_client()
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        temperature=temperature,
        max_tokens=6000,
    )
    text = response.choices[0].message.content.strip()
    st.session_state["raw_ai_output"] = text 
    
    # PERBAIKAN: Pembersihan lebih agresif
    # Menghapus semua teks sebelum { dan setelah }
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1:
        text = text[start:end+1]
        
    # Menghapus karakter kontrol yang sering merusak JSON
    text = text.replace('\n', ' ').replace('\r', '')
    
    return json.loads(text)

# ==============================================================================
# PROMPT CHAINING SCHEMAS
# ==============================================================================

def prompt_step_1(form: dict) -> str:
    return f"""Kamu adalah pakar Kurikulum Merdeka Deep learning Berbasis Cinta di Indonesia. Buat Bagian A (Identifikasi) dan C (Desain Pembelajaran) 
untuk Mapel: {form['mapel']}, Jenjang: {form['kelas']}, Topik: {form['bab']}.
Fokuskan materi secara eksklusif pada Kurikulum Merdeka Deep Learning Berbasis Cinta 5 pilar (KBC).
PENTING: Pada bagian "capaian_pembelajaran", WAJIB merujuk dan menyebutkan secara eksplisit sesuai "KMA Nomor 1503 Tahun 2025".

Balas HANYA dengan JSON valid sesuai skema berikut:
{{
  "identifikasi": {{
    "pengetahuan_awal": ["string", "string"],
    "minat_belajar": ["string", "string"],
    "latar_belakang": "paragraf singkat",
    "kebutuhan_belajar": ["string", "string"],
    "dimensi_profil": ["string", "string"],
    "panca_cinta": ["string", "string"]
  }},
  "desain": {{
    "capaian_pembelajaran": "paragraf yang menyebutkan KMA Nomor 1503 Tahun 2025",
    "tujuan_pembelajaran": ["string", "string"],
    "lintas_disiplin": ["string", "string"]
  }}
}}"""

def prompt_step_2(form: dict, step1_data: dict) -> str:
    n = form["jumlah_pertemuan"]
    return f"""Melanjutkan modul {form['mapel']} {form['kelas']} bab {form['bab']}.
Buat Pengalaman Belajar untuk TEPAT {n} pertemuan secara logis berurutan.
Fokuskan secara utuh pada Kurikulum Merdeka Deep Learning Berbasis Cinta 5 pilar (KBC).
Integrasikan prinsip Deep Learning (Mindful, Meaningful, Joyful) secara jelas ke dalam kalimat kegiatannya.

Balas HANYA dengan JSON valid sesuai skema berikut:
{{
  "pertemuan": [
    {{
      "nomor": 1,
      "tujuan": "string",
      "pembuka": ["string", "string"],
      "inti": ["string", "string", "string"],
      "penutup": ["string", "string"]
    }}
  ]
}}"""

def prompt_step_3(form: dict, step2_data: dict) -> str:
    return f"""Tahap akhir penyusunan modul {form['mapel']} bab {form['bab']}. 
Buat detail Lampiran berdasarkan pertemuan yang telah disusun.
Fokuskan pada pendekatan Kurikulum Merdeka Deep Learning Berbasis Cinta 5 Pilar.

Balas HANYA dengan JSON valid sesuai skema berikut:
{{
  "asesmen_awal": ["string", "string"],
  "asesmen_sumatif": ["string", "string", "string", "string", "string"],
  "materi_ajar": "paragraf panjang / ringkasan materi",
  "lkpd": [
    {{
      "nomor": 1,
      "judul": "string",
      "instruksi": ["string", "string"]
    }}
  ],
  "remedial": "paragraf singkat program remedial",
  "pengayaan": "paragraf singkat program pengayaan",
  "refleksi": ["string", "string", "string"]
}}"""

# ==============================================================================
# DOCX BUILDER 
# ==============================================================================

def set_cell_background(cell, hex_color: str):
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
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    table.columns[0].width = Cm(17)
    cell = table.rows[0].cells[0]
    set_cell_background(cell, hex_color)
    style_cell_text(cell, text, bold=True, color_hex="FFFFFF", size=size)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table

def field_table(doc):
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(4.5)
    table.columns[1].width = Cm(12.5)
    return table

def doc_bullet_style(cell):
    return cell.paragraphs[0].style

def add_field_row(table, label, content_items):
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

def build_docx(form: dict, full_data: dict) -> bytes:
    doc = Document()
    doc.sections[0].left_margin = Cm(2)
    doc.sections[0].right_margin = Cm(2)
    doc.sections[0].page_width = Cm(21)

    d1, d2, d3 = full_data["step1"], full_data["step2"], full_data["step3"]

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
        ("Jenjang / Kelas", form["kelas"]),
        ("Semester", form["semester"]),
        ("Alokasi Waktu", f"{len(d2['pertemuan'])} Pertemuan x {form['alokasi']}"),
        ("Bab / Topik", form["bab"]),
        ("Tahun Pelajaran", form["tahun_pelajaran"]),
        ("Penyusun", form["penyusun"]),
        ("Sekolah", form["sekolah"]),
    ]:
        add_field_row(identity, label, value)
    doc.add_paragraph()

    # --- A. Identifikasi Peserta Didik ---
    banner(doc, "A. IDENTIFIKASI PESERTA DIDIK", COLOR_SECTION_A)
    ident = field_table(doc)
    add_field_row(ident, "Pengetahuan Awal", d1["identifikasi"]["pengetahuan_awal"])
    add_field_row(ident, "Minat Belajar", d1["identifikasi"]["minat_belajar"])
    add_field_row(ident, "Latar Belakang", d1["identifikasi"]["latar_belakang"])
    add_field_row(ident, "Kebutuhan Belajar", d1["identifikasi"]["kebutuhan_belajar"])
    add_field_row(ident, "Dimensi Profil Kelulusan", d1["identifikasi"]["dimensi_profil"])
    add_field_row(ident, "Topik Panca Cinta", d1["identifikasi"]["panca_cinta"])
    doc.add_paragraph()

    # --- C. Desain Pembelajaran ---
    banner(doc, "C. DESAIN PEMBELAJARAN", COLOR_SECTION_B)
    desain = field_table(doc)
    add_field_row(desain, "Capaian Pembelajaran (CP)", d1["desain"]["capaian_pembelajaran"])
    add_field_row(desain, "Tujuan Pembelajaran (TP)", d1["desain"]["tujuan_pembelajaran"])
    add_field_row(desain, "Lintas Disiplin Ilmu", d1["desain"]["lintas_disiplin"])
    doc.add_paragraph()

    # --- Pengalaman Belajar per pertemuan ---
    for p in d2["pertemuan"]:
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
    add_field_row(asesmen_awal, "Pertanyaan Lisan", d3["asesmen_awal"])
    banner(doc, "B. ASESMEN SUMATIF \u2013 SOAL HOTS", COLOR_MEETING, size=10.5)
    sumatif = field_table(doc)
    add_field_row(sumatif, "Soal Uraian", [f"{i}. {s}" for i, s in enumerate(d3["asesmen_sumatif"], 1)])
    doc.add_paragraph()

    # --- Lampiran II: Materi Ajar ---
    banner(doc, "LAMPIRAN II \u2013 MATERI AJAR", COLOR_LAMPIRAN_II)
    doc.add_paragraph(d3["materi_ajar"])
    doc.add_paragraph()

    # --- Lampiran III: LKPD ---
    banner(doc, "LAMPIRAN III \u2013 LKPD (LEMBAR KERJA PESERTA DIDIK)", COLOR_LAMPIRAN_III)
    for p in d3["lkpd"]:
        banner(doc, f"LKPD PERTEMUAN {p['nomor']} \u2013 {p.get('judul', 'Tugas')}", COLOR_LAMPIRAN_III, size=10.5)
        lkpd = field_table(doc)
        add_field_row(lkpd, "Instruksi", p["instruksi"])
        doc.add_paragraph()

    # --- Lampiran V: Tindak Lanjut dan Refleksi ---
    banner(doc, "LAMPIRAN V \u2013 TINDAK LANJUT DAN REFLEKSI", COLOR_LAMPIRAN_V)
    banner(doc, "A. PROGRAM REMEDIAL", COLOR_REMEDIAL, size=10.5)
    doc.add_paragraph(d3["remedial"])
    banner(doc, "B. PROGRAM PENGAYAAN", COLOR_PENGAYAAN, size=10.5)
    doc.add_paragraph(d3["pengayaan"])
    banner(doc, "C. INSTRUMEN REFLEKSI PEMBELAJARAN", COLOR_REFLEKSI, size=10.5)
    refleksi = field_table(doc)
    add_field_row(refleksi, "Pertanyaan Refleksi", d3["refleksi"])
    
    # --- TANDA TANGAN (KOLOM PENGESAHAN) ---
    doc.add_paragraph() 
    doc.add_paragraph()
    
    # Membuat tabel transparan (tanpa garis) untuk tanda tangan bersebelahan
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.columns[0].width = Cm(8.5)
    sig_table.columns[1].width = Cm(8.5)
    
    # Sel kiri: Kepala Madrasah
    p_kepsek = sig_table.rows[0].cells[0].paragraphs[0]
    p_kepsek.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kepsek.add_run("\nMengetahui,\nKepala Madrasah\n\n\n\n")
    run_kepsek = p_kepsek.add_run(form['kepala_madrasah'])
    run_kepsek.bold = True
    
    # Sel kanan: Guru Kelas (beserta titimangsa)
    p_guru = sig_table.rows[0].cells[1].paragraphs[0]
    p_guru.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_guru.add_run(f"{form['titimangsa']}\nGuru Mapel / Kelas\n\n\n\n")
    run_guru = p_guru.add_run(form['penyusun'])
    run_guru.bold = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ==============================================================================
# FUNGSI AUTO-DOWNLOAD
# ==============================================================================
def trigger_download(file_bytes, filename):
    b64 = base64.b64encode(file_bytes).decode()
    html = f'''
        <a id="download-link" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}"></a>
        <script>
            document.getElementById("download-link").click();
        </script>
    '''
    components.html(html, height=0)

# ==============================================================================
# UI STREAMLIT
# ==============================================================================
st.title("📘 MI MIFTAHUSSALAM ADMIN GURU (Fokus KBC & Auto Download)")

with st.form("form_modul"):
    st.subheader("Data Modul")
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Pecahan dan Desimal")
        kelas = st.selectbox("Jenjang / Kelas", list(JENJANG_FASE.keys()), index=6) 
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan", min_value=1, max_value=8, value=2)
        alokasi = st.text_input("Alokasi Waktu per Pertemuan", value="4 JP x 35 menit")
        sekolah = st.text_input("Sekolah", placeholder="MI Miftahussalam")
        tahun_pelajaran = st.text_input("Tahun Pelajaran", value="2026/2027")

    st.divider()
    st.subheader("Data Pengesahan (Tanda Tangan)")
    col3, col4 = st.columns(2)
    
    # Logika penanggalan dinamis (Otomatis hari ini)
    now = datetime.datetime.now()
    bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    nama_bulan = bulan_indo[now.month - 1]
    titimangsa_otomatis = f"Bogor, {now.day} {nama_bulan} {now.year}"

    with col3:
        titimangsa = st.text_input("Titimangsa (Tempat, Waktu)", value=titimangsa_otomatis)
        penyusun = st.text_input("Penyusun (Nama Guru)", placeholder="Erian Kurniawan")
    with col4:
        kepala_madrasah = st.text_input("Nama Kepala Madrasah", placeholder="Drs. Andi Supriadi")

    submitted = st.form_submit_button("✨ Buat Modul Ajar (Butuh ~1 Menit)", use_container_width=True)

if submitted:
    if not (mapel and bab and penyusun and sekolah and kepala_madrasah):
        st.warning("Lengkapi minimal Mata Pelajaran, Bab/Topik, Penyusun, Kepala Madrasah, dan Sekolah.")
    else:
        form = dict(
            mapel=mapel, bab=bab, kelas=kelas, semester=semester,
            jumlah_pertemuan=int(jumlah_pertemuan), alokasi=alokasi,
            penyusun=penyusun, sekolah=sekolah, tahun_pelajaran=tahun_pelajaran,
            titimangsa=titimangsa, kepala_madrasah=kepala_madrasah
        )
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.write("⏳ Langkah 1/3: Menyusun Identitas, CP (KMA 1503/2025) & Desain Pembelajaran...")
            d1 = call_ai(prompt_step_1(form))
            progress_bar.progress(33)
            
            status_text.write("⏳ Langkah 2/3: Merancang Pengalaman Belajar (Mindful, Meaningful, Joyful)...")
            d2 = call_ai(prompt_step_2(form, d1))
            progress_bar.progress(66)
            
            status_text.write("⏳ Langkah 3/3: Menyiapkan Asesmen & LKPD...")
            d3 = call_ai(prompt_step_3(form, d2))
            progress_bar.progress(100)
            
            status_text.success("✅ Modul Ajar Berhasil Disusun! Memunculkan pop-up unduhan otomatis...")
            full_data = {"step1": d1, "step2": d2, "step3": d3}
            st.session_state["full_data"] = full_data
            st.session_state["form"] = form
            
            # --- Generate DOCX dan Eksekusi Auto Download ---
            docx_bytes = build_docx(form, full_data)
            safe_mapel = re.sub(r'[^a-zA-Z0-9_\-]', '_', form['mapel'])
            safe_kelas = re.sub(r'[^a-zA-Z0-9_\-]', '_', form['kelas'].split()[0])
            filename = f"Modul_Ajar_KBC_{safe_mapel}_{safe_kelas}.docx"
            
            trigger_download(docx_bytes, filename)
            
        except json.JSONDecodeError as e:
            st.error("AI gagal menghasilkan JSON yang valid pada salah satu tahap.")
            with st.expander("🔍 Lihat Hasil Mentah (Raw Output) AI untuk Debugging"):
                st.code(st.session_state.get("raw_ai_output", ""))
        except Exception as e:
            st.error(f"Terjadi kesalahan koneksi/kode: {e}")

if "full_data" in st.session_state:
    form = st.session_state["form"]
    full_data = st.session_state["full_data"]
    
    st.divider()
    
    safe_mapel = re.sub(r'[^a-zA-Z0-9_\-]', '_', form['mapel'])
    safe_kelas = re.sub(r'[^a-zA-Z0-9_\-]', '_', form['kelas'].split()[0])
    
    st.subheader(f"🎉 Selesai! Pratinjau & Unduh Modul: {form['mapel']} \u2013 {form['bab']}")
    
    # Tombol unduh cadangan (fallback jika auto-download diblokir browser)
    docx_bytes = build_docx(form, full_data)
    st.download_button(
        "⬇️ Unduh Ulang sebagai Word (.docx)",
        data=docx_bytes,
        file_name=f"Modul_Ajar_KBC_{safe_mapel}_{safe_kelas}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
