"""
Generator Dokumen Admin Guru MI (KBC & KMA 1503/2025)
--------------------------------------------------------------------------------
Fitur: Modul Ajar, CP & ATP, Prota, Promes, Jurnal (Support Landscape)
Fokus: Kurikulum Merdeka Deep Learning Berbasis Cinta (KBC)
"""

import io
import json
import re
import datetime
import base64
import time

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
# PALET WARNA & KONFIGURASI
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

st.set_page_config(page_title="MIFSAL ADMIN GURU V4", page_icon="📘", layout="wide")

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
    
    # Pembersihan agresif untuk menangkap JSON
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1:
        text = text[start:end+1]
        
    text = text.replace('\n', ' ').replace('\r', '')
    return json.loads(text)

# ==============================================================================
# PROMPT SCHEMAS (AI)
# ==============================================================================

# 1. Prompts khusus Modul Ajar (3 Steps)
def prompt_step_1(form: dict) -> str:
    return f"""Kamu pakar Kurikulum Merdeka Deep Learning Berbasis Cinta (KBC). Buat Bagian A & C modul
untuk Mapel: {form['mapel']}, Jenjang: {form['kelas']}, Topik: {form['bab']}.
PENTING: Pada "capaian_pembelajaran", WAJIB sebutkan eksplisit sesuai "KMA Nomor 1503 Tahun 2025".
Balas HANYA dengan JSON:
{{"identifikasi": {{"pengetahuan_awal": ["string"], "minat_belajar": ["string"], "latar_belakang": "string", "kebutuhan_belajar": ["string"], "dimensi_profil": ["string"], "panca_cinta": ["string"]}}, "desain": {{"capaian_pembelajaran": "paragraf dengan KMA 1503 Tahun 2025", "tujuan_pembelajaran": ["string"], "lintas_disiplin": ["string"]}}}}"""

def prompt_step_2(form: dict, step1_data: dict) -> str:
    return f"""Lanjutkan modul {form['mapel']} {form['kelas']} bab {form['bab']}.
Buat Pengalaman Belajar {form['jumlah_pertemuan']} pertemuan. Integrasikan prinsip KBC (Mindful, Meaningful, Joyful).
Balas HANYA dengan JSON:
{{"pertemuan": [{{"nomor": 1, "tujuan": "string", "pembuka": ["string"], "inti": ["string"], "penutup": ["string"]}}]}}"""

def prompt_step_3(form: dict, step2_data: dict) -> str:
    return f"""Tahap akhir modul {form['mapel']} bab {form['bab']}. Fokus KBC 5 Pilar.
Balas HANYA dengan JSON:
{{"asesmen_awal": ["string"], "asesmen_sumatif": ["string"], "materi_ajar": "string", "lkpd": [{{"nomor": 1, "judul": "string", "instruksi": ["string"]}}], "remedial": "string", "pengayaan": "string", "refleksi": ["string"]}}"""

# 2. Prompt Umum untuk Dokumen Lain (CP/ATP, Prota, Promes, Jurnal)
def prompt_generic_doc(form: dict, tipe_dokumen: str) -> str:
    return f"""Kamu pakar Kurikulum Merdeka KBC. Buat isi konten untuk dokumen '{tipe_dokumen}' 
Mata Pelajaran: {form['mapel']}, Kelas: {form['kelas']}, Topik: {form['bab']}.
Gunakan format KMA 1503 Tahun 2025.
Balas HANYA dengan JSON valid berisi rincian poin-poin/materi:
{{"judul_dokumen": "string", "deskripsi_umum": "string", "rincian_tabel": [{{"kolom_1": "string (Topik/Waktu)", "kolom_2": "string (Deskripsi/CP/Kegiatan)"}}]}}"""


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

def doc_bullet_style(cell):
    return cell.paragraphs[0].style

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
            p.add_run(f"\u2022 {item}").font.size = Pt(10.5)
            first = False
    else:
        value_cell.paragraphs[0].add_run(str(content_items)).font.size = Pt(10.5)
    return row

def add_signatures(doc, form):
    doc.add_paragraph() 
    doc.add_paragraph()
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.columns[0].width = Cm(8.5)
    sig_table.columns[1].width = Cm(8.5)
    
    p_kepsek = sig_table.rows[0].cells[0].paragraphs[0]
    p_kepsek.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kepsek.add_run("\nMengetahui,\nKepala Madrasah\n\n\n\n")
    p_kepsek.add_run(form['kepala_madrasah']).bold = True
    
    p_guru = sig_table.rows[0].cells[1].paragraphs[0]
    p_guru.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_guru.add_run(f"{form['titimangsa']}\nGuru Mapel / Kelas\n\n\n\n")
    p_guru.add_run(form['penyusun']).bold = True

def create_base_doc(landscape=False):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    
    if landscape:
        new_width, new_height = section.page_height, section.page_width
        section.orientation = 1 # LANDSCAPE
        section.page_width = new_width
        section.page_height = new_height
    else:
        section.page_width = Cm(21)
        
    return doc

# Builder Spesifik: Modul Ajar
def build_modul_ajar(form: dict, full_data: dict) -> bytes:
    doc = create_base_doc(landscape=False)
    d1, d2, d3 = full_data["step1"], full_data["step2"], full_data["step3"]

    title_table = doc.add_table(rows=1, cols=1)
    cell = title_table.rows[0].cells[0]
    set_cell_background(cell, COLOR_TITLE)
    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run("MODUL AJAR").bold = True
    p1.runs[0].font.size, p1.runs[0].font.color.rgb = Pt(16), RGBColor.from_string("FFFFFF")
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run("KURIKULUM BERBASIS CINTA \u2013 PENDEKATAN DEEP LEARNING").bold = True
    p2.runs[0].font.size, p2.runs[0].font.color.rgb = Pt(11), RGBColor.from_string("FFFFFF")
    doc.add_paragraph()

    banner(doc, "IDENTITAS MODUL AJAR", COLOR_IDENTITY_HEAD)
    identity = field_table(doc)
    for label, value in [
        ("Mata Pelajaran", form["mapel"]), ("Jenjang / Kelas", form["kelas"]),
        ("Semester", form["semester"]), ("Alokasi Waktu", f"{len(d2['pertemuan'])} Pertemuan x {form['alokasi']}"),
        ("Bab / Topik", form["bab"]), ("Tahun Pelajaran", form["tahun_pelajaran"]),
        ("Penyusun", form["penyusun"]), ("Sekolah", form["sekolah"]),
    ]:
        add_field_row(identity, label, value)
    doc.add_paragraph()

    banner(doc, "A. IDENTIFIKASI PESERTA DIDIK", COLOR_SECTION_A)
    ident = field_table(doc)
    add_field_row(ident, "Pengetahuan Awal", d1["identifikasi"]["pengetahuan_awal"])
    add_field_row(ident, "Minat Belajar", d1["identifikasi"]["minat_belajar"])
    add_field_row(ident, "Latar Belakang", d1["identifikasi"]["latar_belakang"])
    add_field_row(ident, "Kebutuhan Belajar", d1["identifikasi"]["kebutuhan_belajar"])
    add_field_row(ident, "Dimensi Profil", d1["identifikasi"]["dimensi_profil"])
    add_field_row(ident, "Topik Panca Cinta", d1["identifikasi"]["panca_cinta"])
    doc.add_paragraph()

    banner(doc, "C. DESAIN PEMBELAJARAN", COLOR_SECTION_B)
    desain = field_table(doc)
    add_field_row(desain, "Capaian Pembelajaran (CP)", d1["desain"]["capaian_pembelajaran"])
    add_field_row(desain, "Tujuan Pembelajaran (TP)", d1["desain"]["tujuan_pembelajaran"])
    add_field_row(desain, "Lintas Disiplin Ilmu", d1["desain"]["lintas_disiplin"])
    doc.add_paragraph()

    for p in d2["pertemuan"]:
        banner(doc, f"PENGALAMAN BELAJAR \u2013 PERTEMUAN {p['nomor']}", COLOR_MEETING, size=11)
        meeting = field_table(doc)
        add_field_row(meeting, "Tujuan Pertemuan", p["tujuan"])
        add_field_row(meeting, "Kegiatan Pembuka", p["pembuka"])
        add_field_row(meeting, "Kegiatan Inti", p["inti"])
        add_field_row(meeting, "Kegiatan Penutup", p["penutup"])
        doc.add_paragraph()

    banner(doc, "LAMPIRAN", COLOR_LAMPIRAN_I)
    asesmen_awal = field_table(doc)
    add_field_row(asesmen_awal, "Asesmen Awal", d3["asesmen_awal"])
    sumatif = field_table(doc)
    add_field_row(sumatif, "Sumatif (HOTS)", [f"- {s}" for s in d3["asesmen_sumatif"]])
    doc.add_paragraph(d3["materi_ajar"])
    
    for p in d3["lkpd"]:
        banner(doc, f"LKPD PERTEMUAN {p['nomor']} \u2013 {p.get('judul', 'Tugas')}", COLOR_LAMPIRAN_III, size=10.5)
        lkpd = field_table(doc)
        add_field_row(lkpd, "Instruksi", p["instruksi"])
        doc.add_paragraph()

    add_signatures(doc, form)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# Builder Spesifik: Dokumen Lain (Landscape otomatis untuk Prota/Promes)
def build_generic_doc(form: dict, ai_data: dict, tipe: str) -> bytes:
    is_landscape = (tipe == "Prota" or tipe == "Promes")
    doc = create_base_doc(landscape=is_landscape)
    
    doc.add_heading(f"DOKUMEN {tipe.upper()}", level=1)
    doc.add_paragraph(f"Mata Pelajaran: {form['mapel']} | Kelas: {form['kelas']}")
    doc.add_paragraph(f"Tahun Pelajaran: {form['tahun_pelajaran']} | Sekolah: {form['sekolah']}")
    doc.add_paragraph()
    
    doc.add_heading(ai_data.get("judul_dokumen", tipe), level=2)
    doc.add_paragraph(ai_data.get("deskripsi_umum", ""))
    doc.add_paragraph()
    
    # Buat Tabel dari rincian AI
    rincian = ai_data.get("rincian_tabel", [])
    if len(rincian) > 0:
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Topik / Alokasi Waktu"
        hdr_cells[1].text = "Deskripsi / Kegiatan Pembelajaran"
        
        for item in rincian:
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get("kolom_1", ""))
            row_cells[1].text = str(item.get("kolom_2", ""))
            
    add_signatures(doc, form)
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
        <a id="dl-{filename}" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}"></a>
        <script>
            document.getElementById("dl-{filename}").click();
        </script>
    '''
    components.html(html, height=0)

# ==============================================================================
# UI STREAMLIT
# ==============================================================================
st.title("📘 MI MIFTAHUSSALAM: ADMIN GURU GENERATOR")
st.markdown("*Kurikulum Merdeka Deep Learning Berbasis Cinta (KBC)*")

with st.form("form_modul"):
    st.subheader("1. Isi Data Administratif")
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Pecahan dan Desimal")
        kelas = st.selectbox("Jenjang / Kelas", list(JENJANG_FASE.keys()), index=6) 
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan (Untuk Modul)", min_value=1, max_value=8, value=2)
        alokasi = st.text_input("Alokasi Waktu per Pertemuan", value="4 JP x 35 menit")
        sekolah = st.text_input("Sekolah", value="MI Miftahussalam")
        tahun_pelajaran = st.text_input("Tahun Pelajaran", value="2026/2027")

    st.divider()
    st.subheader("2. Data Pengesahan")
    col3, col4 = st.columns(2)
    now = datetime.datetime.now()
    bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    titimangsa_otomatis = f"Bogor, {now.day} {bulan_indo[now.month - 1]} {now.year}"

    with col3:
        titimangsa = st.text_input("Titimangsa (Tempat, Waktu)", value=titimangsa_otomatis)
        penyusun = st.text_input("Penyusun (Nama Guru)", placeholder="Erian Kurniawan")
    with col4:
        kepala_madrasah = st.text_input("Nama Kepala Madrasah", placeholder="Drs. Andi Supriadi")

    st.divider()
    st.subheader("3. Pilih Dokumen yang Dibuat")
    pilihan_dokumen = st.multiselect(
        "Pilih satu atau lebih dokumen untuk di-generate otomatis:",
        ["Modul Ajar", "CP & ATP", "Prota", "Promes", "Jurnal Mengajar"],
        default=["Modul Ajar"]
    )

    submitted = st.form_submit_button("✨ Eksekusi & Generate (Auto-Download)", use_container_width=True)

if submitted:
    if not (mapel and bab and penyusun and sekolah and kepala_madrasah):
        st.warning("Lengkapi minimal Mata Pelajaran, Bab/Topik, Penyusun, Kepala, dan Sekolah.")
    elif not pilihan_dokumen:
        st.warning("Silakan pilih minimal 1 dokumen yang ingin dibuat pada menu dropdown.")
    else:
        form = dict(
            mapel=mapel, bab=bab, kelas=kelas, semester=semester,
            jumlah_pertemuan=int(jumlah_pertemuan), alokasi=alokasi,
            penyusun=penyusun, sekolah=sekolah, tahun_pelajaran=tahun_pelajaran,
            titimangsa=titimangsa, kepala_madrasah=kepala_madrasah
        )
        
        safe_mapel = re.sub(r'[^a-zA-Z0-9_\-]', '_', form['mapel'])
        safe_kelas = re.sub(r'[^a-zA-Z0-9_\-]', '_', form['kelas'].split()[0])
        
        st.session_state["hasil_generate"] = {}
        
        # PROSES BERDASARKAN PILIHAN
        try:
            if "Modul Ajar" in pilihan_dokumen:
                st.write("📘 **Memproses Modul Ajar (3 Tahap)...**")
                bar = st.progress(0)
                d1 = call_ai(prompt_step_1(form)); bar.progress(33)
                d2 = call_ai(prompt_step_2(form, d1)); bar.progress(66)
                d3 = call_ai(prompt_step_3(form, d2)); bar.progress(100)
                
                doc_bytes = build_modul_ajar(form, {"step1": d1, "step2": d2, "step3": d3})
                filename = f"Modul_Ajar_KBC_{safe_mapel}_{safe_kelas}.docx"
                st.session_state["hasil_generate"][filename] = doc_bytes
                trigger_download(doc_bytes, filename)
                time.sleep(1) # Jeda antar pop-up agar browser tidak bingung
                
            for tipe in ["CP & ATP", "Prota", "Promes", "Jurnal Mengajar"]:
                if tipe in pilihan_dokumen:
                    st.write(f"📄 **Memproses {tipe}...**")
                    ai_data = call_ai(prompt_generic_doc(form, tipe))
                    doc_bytes = build_generic_doc(form, ai_data, tipe)
                    
                    safe_tipe = tipe.replace(" & ", "_").replace(" ", "_")
                    filename = f"{safe_tipe}_{safe_mapel}_{safe_kelas}.docx"
                    st.session_state["hasil_generate"][filename] = doc_bytes
                    trigger_download(doc_bytes, filename)
                    time.sleep(1)
            
            st.success("✅ Semua dokumen berhasil disusun dan sedang diunduh!")
            
        except json.JSONDecodeError as e:
            st.error("AI gagal menghasilkan format data yang valid. Coba ulangi atau kurangi jumlah dokumen.")
            with st.expander("🔍 Lihat Hasil Mentah AI"):
                st.code(st.session_state.get("raw_ai_output", ""))
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

# Tampilkan tombol unduh cadangan
if "hasil_generate" in st.session_state and st.session_state["hasil_generate"]:
    st.divider()
    st.write("⬇️ **Tombol Unduh Manual (Jika pop-up terblokir):**")
    for fname, fbytes in st.session_state["hasil_generate"].items():
        st.download_button(label=f"Unduh {fname}", data=fbytes, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
