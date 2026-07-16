"""
Generator Dokumen Admin Guru MI (KBC & KMA 1503/2025)
--------------------------------------------------------------------------------
Pembaruan: Format Tabel Standar Baku Indonesia untuk Prota, Promes, ATP, dan Jurnal.
(Mendukung penggabungan sel/merge cells & kolom mingguan pada Promes)
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
# PALET WARNA
# ==============================================================================
COLOR_TITLE = "1F4E79"       
COLOR_IDENTITY_HEAD = "2E75B6"   
COLOR_LABEL = "DEEAF1"       
COLOR_VALUE = "FFFFFF"       
COLOR_MEETING = "843C0C"     
COLOR_LAMPIRAN = "375623"

MODEL_NAME = "nvidia/nemotron-3-ultra-550b-a55b"
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"

JENJANG_FASE = {
    "RA/TK (Fase Fondasi)": "Fondasi", "Kelas 1 SD/MI (Fase A)": "A", "Kelas 2 SD/MI (Fase A)": "A",
    "Kelas 3 SD/MI (Fase B)": "B", "Kelas 4 SD/MI (Fase B)": "B", "Kelas 5 SD/MI (Fase C)": "C",
    "Kelas 6 SD/MI (Fase C)": "C", "Kelas 7 SMP/MTs (Fase D)": "D", "Kelas 8 SMP/MTs (Fase D)": "D",
    "Kelas 9 SMP/MTs (Fase D)": "D", "Kelas 10 SMA/MA (Fase E)": "E", "Kelas 11 SMA/MA (Fase F)": "F",
    "Kelas 12 SMA/MA (Fase F)": "F",
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
        model=MODEL_NAME, messages=[{"role": "user", "content": prompt}],
        temperature=temperature, max_tokens=6000,
    )
    text = response.choices[0].message.content.strip()
    st.session_state["raw_ai_output"] = text 
    
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1:
        text = text[start:end+1]
        
    text = text.replace('\n', ' ').replace('\r', '')
    return json.loads(text)

# ==============================================================================
# PROMPT SCHEMAS (KHUSUS FORMAT STANDAR)
# ==============================================================================

# Prompt Modul Ajar (Langkah 1-3)
def prompt_step_1(form):
    return f"""Buat Bagian A & C modul KBC Mapel: {form['mapel']}, Kelas: {form['kelas']}, Topik: {form['bab']}. CP WAJIB sebutkan "KMA Nomor 1503 Tahun 2025". Balas HANYA JSON: {{"identifikasi": {{"pengetahuan_awal": ["str"], "minat_belajar": ["str"], "latar_belakang": "str", "kebutuhan_belajar": ["str"], "dimensi_profil": ["str"], "panca_cinta": ["str"]}}, "desain": {{"capaian_pembelajaran": "str", "tujuan_pembelajaran": ["str"], "lintas_disiplin": ["str"]}}}}"""

def prompt_step_2(form, step1):
    return f"""Lanjutkan modul {form['mapel']} {form['kelas']} bab {form['bab']}. Buat Pengalaman Belajar {form['jumlah_pertemuan']} pertemuan (Mindful, Meaningful, Joyful). Balas HANYA JSON: {{"pertemuan": [{{"nomor": 1, "tujuan": "str", "pembuka": ["str"], "inti": ["str"], "penutup": ["str"]}}]}}"""

def prompt_step_3(form, step2):
    return f"""Tahap akhir modul {form['mapel']} bab {form['bab']}. Balas HANYA JSON: {{"asesmen_awal": ["str"], "asesmen_sumatif": ["str"], "materi_ajar": "str", "lkpd": [{{"nomor": 1, "judul": "str", "instruksi": ["str"]}}], "remedial": "str", "pengayaan": "str", "refleksi": ["str"]}}"""

# Prompts Dokumen Administratif
def prompt_cpatp(form):
    return f"""Buat isi CP dan ATP Mapel {form['mapel']} {form['kelas']} Topik {form['bab']} berdasar KMA 1503/2025. Balas HANYA JSON: {{"rows": [{{"elemen": "str", "cp": "str", "tp": "str", "atp": "str", "jp": "str"}}]}}"""

def prompt_prota(form):
    return f"""Buat isi Program Tahunan Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Bagikan ke beberapa baris elemen. Balas HANYA JSON: {{"rows": [{{"no": "1", "elemen_cp": "str", "topik_tp": "str", "jp": "str", "semester": "1"}}]}}"""

def prompt_promes(form):
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    return f"""Buat rincian Program Semester Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Pecah materi ke dalam bulan {bulan}. "minggu" berisi array angka minggu (1-5) pelaksanaannya. Balas HANYA JSON: {{"rows": [{{"no": "1", "materi_tp": "str", "jp": "str", "bulan": "Juli", "minggu": [1, 2]}}]}}"""

def prompt_jurnal(form):
    return f"""Buat isi Jurnal Mengajar Harian untuk {form['jumlah_pertemuan']} pertemuan. Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Balas HANYA JSON: {{"rows": [{{"pertemuan": "1", "topik": "str", "aktivitas": "str (Deep Learning)", "asesmen": "str"}}]}}"""


# ==============================================================================
# FUNGSI PEMBANTU FORMATTING DOCX
# ==============================================================================

def set_cell_background(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)

def style_cell(cell, text, bold=False, color="000000", center=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)

def create_base_doc(landscape=False):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(1.5); section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5); section.bottom_margin = Cm(1.5)
    if landscape:
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width
    return doc

def add_signatures(doc, form):
    doc.add_paragraph("\n")
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.columns[0].width = Cm(8.5); sig_table.columns[1].width = Cm(8.5)
    p1 = sig_table.rows[0].cells[0].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run("\nMengetahui,\nKepala Madrasah\n\n\n\n")
    p1.add_run(form['kepala_madrasah']).bold = True
    p2 = sig_table.rows[0].cells[1].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"{form['titimangsa']}\nGuru Mata Pelajaran\n\n\n\n")
    p2.add_run(form['penyusun']).bold = True

def create_header(doc, title, form):
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Mata Pelajaran : {form['mapel']} \t\t Semester : {form['semester']}")
    doc.add_paragraph(f"Kelas / Fase   : {form['kelas']} \t\t Tahun Pelajaran : {form['tahun_pelajaran']}")
    doc.add_paragraph()

# ==============================================================================
# BUILDERS DOKUMEN KHUSUS (STANDAR INDONESIA)
# ==============================================================================

# 1. BUILDER MODUL AJAR (Tetap memakai format sebelumnya)
def build_modul_ajar(form, full_data):
    doc = create_base_doc(landscape=False)
    # [Logika Modul Ajar dihilangkan sementara dari visual ini untuk menghemat teks,
    # Namun akan terisi persis seperti kode yang sudah berjalan sebelumnya. 
    # Karena fokus kita merevisi 4 dokumen di bawah ini]
    doc.add_heading("MODUL AJAR KBC", level=1)
    doc.add_paragraph("Silakan gunakan builder dari kode sebelumnya untuk bagian ini.")
    return b""

# 2. BUILDER CP & ATP
def build_cpatp(form, ai_data):
    doc = create_base_doc(landscape=True)
    create_header(doc, "CAPAIAN PEMBELAJARAN (CP) & ALUR TUJUAN PEMBELAJARAN (ATP)", form)
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Elemen", "Capaian Pembelajaran (CP)", "Tujuan Pembelajaran (TP)", "Alur Tujuan Pembelajaran (ATP)", "JP"]
    
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), COLOR_TITLE)
        style_cell(table.cell(0, i), h, bold=True, color="FFFFFF", center=True)
        
    for row in ai_data.get("rows", []):
        r = table.add_row().cells
        style_cell(r[0], row.get("elemen", ""))
        style_cell(r[1], row.get("cp", ""))
        style_cell(r[2], row.get("tp", ""))
        style_cell(r[3], row.get("atp", ""))
        style_cell(r[4], row.get("jp", ""), center=True)
        
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# 3. BUILDER PROTA (PROGRAM TAHUNAN)
def build_prota(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "PROGRAM TAHUNAN (PROTA)", form)
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["No", "Elemen / CP", "Topik / Tujuan Pembelajaran", "Alokasi Waktu (JP)", "Semester"]
    
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), COLOR_TITLE)
        style_cell(table.cell(0, i), h, bold=True, color="FFFFFF", center=True)
    
    table.columns[0].width = Cm(1.0)
    table.columns[3].width = Cm(2.5)
    table.columns[4].width = Cm(2.0)
        
    for row in ai_data.get("rows", []):
        r = table.add_row().cells
        style_cell(r[0], row.get("no", ""), center=True)
        style_cell(r[1], row.get("elemen_cp", ""))
        style_cell(r[2], row.get("topik_tp", ""))
        style_cell(r[3], row.get("jp", ""), center=True)
        style_cell(r[4], row.get("semester", ""), center=True)
        
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# 4. BUILDER PROMES (PROGRAM SEMESTER - KOMPLEKS)
def build_promes(form, ai_data):
    doc = create_base_doc(landscape=True)
    create_header(doc, "PROGRAM SEMESTER (PROMES)", form)
    
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    
    total_cols = 3 + (len(bulan) * 5) + 1 # No, Materi, JP + (6 bulan x 5 minggu) + Keterangan
    table = doc.add_table(rows=2, cols=total_cols)
    table.style = 'Table Grid'
    
    # Header Baris 1
    table.cell(0, 0).merge(table.cell(1, 0)); style_cell(table.cell(0, 0), "No", bold=True, center=True)
    table.cell(0, 1).merge(table.cell(1, 1)); style_cell(table.cell(0, 1), "Tujuan Pembelajaran / Materi", bold=True, center=True)
    table.cell(0, 2).merge(table.cell(1, 2)); style_cell(table.cell(0, 2), "JP", bold=True, center=True)
    
    # Set warna Header utama
    for i in range(3): set_cell_background(table.cell(0, i), COLOR_LABEL)
    
    col_idx = 3
    for b in bulan:
        # Merge 5 cell ke kanan untuk nama bulan
        table.cell(0, col_idx).merge(table.cell(0, col_idx + 4))
        style_cell(table.cell(0, col_idx), b, bold=True, center=True)
        set_cell_background(table.cell(0, col_idx), COLOR_LABEL)
        
        # Header Baris 2 (Minggu 1-5)
        for w in range(5):
            style_cell(table.cell(1, col_idx + w), str(w + 1), bold=True, center=True)
            set_cell_background(table.cell(1, col_idx + w), "EFEFEF")
        col_idx += 5
        
    table.cell(0, total_cols - 1).merge(table.cell(1, total_cols - 1))
    style_cell(table.cell(0, total_cols - 1), "Ket.", bold=True, center=True)
    set_cell_background(table.cell(0, total_cols - 1), COLOR_LABEL)
    
    # Isi Data dari AI
    for row in ai_data.get("rows", []):
        r = table.add_row().cells
        style_cell(r[0], row.get("no", ""), center=True)
        style_cell(r[1], row.get("materi_tp", ""))
        style_cell(r[2], row.get("jp", ""), center=True)
        
        # Logika centang minggu
        target_bulan = row.get("bulan", "")
        minggu_aktif = row.get("minggu", [])
        
        idx = 3
        for b in bulan:
            for w in range(1, 6):
                if target_bulan.lower() == b.lower() and w in minggu_aktif:
                    set_cell_background(r[idx], COLOR_TITLE) # Beri warna blok pada minggu yg dipilih
                idx += 1
                
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# 5. BUILDER JURNAL MENGAJAR
def build_jurnal(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "JURNAL MENGAJAR HARIAN", form)
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Pertemuan", "Topik / Materi", "Aktivitas Deep Learning", "Asesmen", "Keterangan / Paraf"]
    
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), COLOR_TITLE)
        style_cell(table.cell(0, i), h, bold=True, color="FFFFFF", center=True)
        
    for row in ai_data.get("rows", []):
        r = table.add_row().cells
        style_cell(r[0], row.get("pertemuan", ""), center=True)
        style_cell(r[1], row.get("topik", ""))
        style_cell(r[2], row.get("aktivitas", ""))
        style_cell(r[3], row.get("asesmen", ""))
        style_cell(r[4], "") # Kosong untuk paraf
        
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()


# ==============================================================================
# FUNGSI AUTO-DOWNLOAD
# ==============================================================================
def trigger_download(file_bytes, filename):
    b64 = base64.b64encode(file_bytes).decode()
    html = f'''<a id="dl-{filename}" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}"></a><script>document.getElementById("dl-{filename}").click();</script>'''
    components.html(html, height=0)

# ==============================================================================
# UI STREAMLIT
# ==============================================================================
st.title("📘 MIFSAL ADMIN GURU GENERATOR (Standar KMA)")
st.markdown("*Kurikulum Merdeka Deep Learning Berbasis Cinta (KBC)*")

with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Pecahan dan Desimal")
        kelas = st.selectbox("Jenjang / Kelas", list(JENJANG_FASE.keys()), index=6) 
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan (Untuk Modul & Jurnal)", min_value=1, max_value=8, value=2)
        alokasi = st.text_input("Alokasi Waktu", value="4 JP x 35 menit")
        sekolah = st.text_input("Sekolah", value="MI Miftahussalam")
        tahun_pelajaran = st.text_input("Tahun Pelajaran", value="2026/2027")

    st.divider()
    col3, col4 = st.columns(2)
    now = datetime.datetime.now()
    bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    titimangsa_otomatis = f"Bogor, {now.day} {bulan_indo[now.month - 1]} {now.year}"

    with col3:
        titimangsa = st.text_input("Titimangsa", value=titimangsa_otomatis)
        penyusun = st.text_input("Penyusun (Guru)", placeholder="Erian Kurniawan")
    with col4:
        kepala_madrasah = st.text_input("Kepala Madrasah", placeholder="Drs. Andi Supriadi")

    st.divider()
    pilihan_dokumen = st.multiselect(
        "Pilih dokumen yang ingin di-generate otomatis:",
        ["CP & ATP", "Prota", "Promes", "Jurnal Mengajar"] # Modul Ajar disembunyikan sementara untuk tes ini
    )

    submitted = st.form_submit_button("✨ Eksekusi & Generate (Auto-Download)", use_container_width=True)

if submitted:
    if not (mapel and bab and penyusun and sekolah and kepala_madrasah):
        st.warning("Lengkapi minimal Mata Pelajaran, Bab/Topik, Penyusun, Kepala, dan Sekolah.")
    elif not pilihan_dokumen:
        st.warning("Silakan pilih minimal 1 dokumen yang ingin dibuat.")
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
        
        try:
            for tipe in pilihan_dokumen:
                st.write(f"📄 **Memproses {tipe}...**")
                if tipe == "CP & ATP":
                    ai_data = call_ai(prompt_cpatp(form))
                    doc_bytes = build_cpatp(form, ai_data)
                elif tipe == "Prota":
                    ai_data = call_ai(prompt_prota(form))
                    doc_bytes = build_prota(form, ai_data)
                elif tipe == "Promes":
                    ai_data = call_ai(prompt_promes(form))
                    doc_bytes = build_promes(form, ai_data)
                elif tipe == "Jurnal Mengajar":
                    ai_data = call_ai(prompt_jurnal(form))
                    doc_bytes = build_jurnal(form, ai_data)
                
                safe_tipe = tipe.replace(" & ", "_").replace(" ", "_")
                filename = f"{safe_tipe}_{safe_mapel}_{safe_kelas}.docx"
                st.session_state["hasil_generate"][filename] = doc_bytes
                trigger_download(doc_bytes, filename)
                time.sleep(1.5)
            
            st.success("✅ Semua dokumen berhasil disusun dan sedang diunduh!")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membangun tabel Word: {e}")

if "hasil_generate" in st.session_state and st.session_state["hasil_generate"]:
    st.divider()
    st.write("⬇️ **Tombol Unduh Manual:**")
    for fname, fbytes in st.session_state["hasil_generate"].items():
        st.download_button(label=f"Unduh {fname}", data=fbytes, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
