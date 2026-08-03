"""
Generator Dokumen Admin Guru MI (KBC & KMA 1503/2025)
--------------------------------------------------------------------------------
Pembaruan: Proses Eksekusi Bertahap (Anti-Error), Modul Ajar 3 Tahap Eksplisit,
Pemisahan CP & ATP, dan Sinkronisasi Format Sesuai File Contoh.
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

MODEL_NAME = "google/diffusiongemma-26b-a4b-it"
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"

JENJANG_FASE = {
    "RA/TK (Fase Fondasi)": "Fondasi", "Kelas 1 SD/MI (Fase A)": "A", "Kelas 2 SD/MI (Fase A)": "A",
    "Kelas 3 SD/MI (Fase B)": "B", "Kelas 4 SD/MI (Fase B)": "B", "Kelas 5 SD/MI (Fase C)": "C",
    "Kelas 6 SD/MI (Fase C)": "C", "Kelas 7 SMP/MTs (Fase D)": "D", "Kelas 8 SMP/MTs (Fase D)": "D",
    "Kelas 9 SMP/MTs (Fase D)": "D", "Kelas 10 SMA/MA (Fase E)": "E", "Kelas 11 SMA/MA (Fase F)": "F",
    "Kelas 12 SMA/MA (Fase F)": "F",
}

st.set_page_config(page_title="MIFSAL ADMIN GURU V3", page_icon="📘", layout="wide")

@st.cache_resource
def get_client():
    api_key = st.secrets.get("NVIDIA_API_KEY")
    if not api_key:
        st.error("NVIDIA_API_KEY belum ada di st.secrets.")
        st.stop()
    return OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)

def call_ai(prompt: str, temperature=0.7) -> dict:
    client = get_client()
    response = client.chat.completions.create(
        model=MODEL_NAME, messages=[{"role": "user", "content": prompt}],
        temperature=temperature, max_tokens=12192,
    )
    text = response.choices[0].message.content.strip()
    st.session_state["raw_ai_output"] = text 
    
    text = text.replace("```json", "").replace("```", "").strip()
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1:
        text = text[start:end+1]
        
    text = text.replace('\n', ' ').replace('\r', '')
    text = re.sub(r'[\x00-\x1f]', '', text)
    text = re.sub(r',\s*([}\]])', r'\1', text)
    
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"JSON Parse Error: {e}")
        return {}

# ==============================================================================
# PROMPT SINKRONISASI TOTAL (MASTER CONTEXT)
# ==============================================================================
def get_sinkronisasi_context(d1_context, d2_context):
    sinkron_text = ""
    if d1_context and isinstance(d1_context, dict) and "desain" in d1_context:
        cp = d1_context["desain"].get("capaian_pembelajaran", "")
        tp = d1_context["desain"].get("tujuan_pembelajaran", [])
        if isinstance(tp, list): tp = ", ".join(tp)
        sinkron_text += f"1. Capaian Pembelajaran (CP): '{cp}'\n2. Tujuan Pembelajaran (TP): '{tp}'\n"

    if d2_context and isinstance(d2_context, dict) and "pertemuan" in d2_context:
        materi_list = [f"Pertemuan {p.get('nomor', '')}: {p.get('materi', '')}" for p in d2_context.get("pertemuan", []) if isinstance(p, dict)]
        materi_str = " | ".join(materi_list)
        sinkron_text += f"3. Rincian Materi Pertemuan WAJIB menggunakan ini: {materi_str}\n"

    if sinkron_text:
        return f"\n\n[SANGAT PENTING - SINKRONISASI DOKUMEN]\nKamu WAJIB mematuhi data dasar berikut agar semua dokumen selaras:\n{sinkron_text}\nJANGAN membuat materi atau TP di luar data di atas!"
    return ""

# ==============================================================================
# PROMPT MODUL AJAR (TIDAK DIRUBAH)
# ==============================================================================
def prompt_step_1(form):
    return f"""Kamu pakar Kurikulum Merdeka Pendekatan Deep Learning Berbasis Cinta dengan 5 pilar (KBC). WAJIB JANGAN SAMPAI BUAT KESALAHAN SEDIKITPUN. Buat Bagian A & B modul gunakan bahasa yang humanis agar tidak terlihat AI, buat semua file menjadi lengkap, kompleks dan sempurna.
untuk Mapel: {form['mapel']}, Jenjang: {form['kelas']}, Topik: {form['bab']}. PENTING: CP dan TP WAJIB mengacu pada "KMA Nomor 1503 Tahun 2025" Untuk pemanfaata Digital Isi Minimal 3, Wajib Masukan semua Panca Cinta Cinta kepada Tuhan (Allah dan Rasul-Nya),Cinta kepada Diri dan Sesama, Cinta kepada Ilmu Pengetahuan, Cinta Lingkungan, Cinta Tanah Air.
untuk Dimensi Profile Kelulusan masukan minimal 4 serta penjelasannya.
PENTING: Balas HANYA dengan JSON valid. DILARANG menggunakan tanda kutip ganda (") di dalam teks string.
Balas HANYA JSON:
{{"identifikasi": {{"pengetahuan_awal": ["str"], "minat_belajar": ["str"], "latar_belakang": "str", "kebutuhan_belajar": ["str"], "dimensi_profil": ["str"], "panca_cinta": ["str"]}}, "desain": {{"capaian_pembelajaran": "str", "tujuan_pembelajaran": ["str"], "lintas_disiplin": ["str"], "topik_pembelajaran": ["str"], "praktik_pedagogi": ["str"], "lingkungan_belajar": ["str"], "kemitraan_pembelajaran": ["str"], "pemanfaatan_digital": ["str"]}}}}"""

def prompt_step_2(form, step1):
    return f"""Lanjutkan modul {form['mapel']} {form['kelas']} bab {form['bab']}. Buat Pengalaman Belajar untuk TEPAT {form['jumlah_pertemuan']} pertemuan. Format 4 elemen untuk setiap kegiatan: fase, aktivitas, waktu, dl.
PENTING: Balas HANYA dengan JSON valid. DILARANG menggunakan tanda kutip ganda (") di dalam teks string.
Balas HANYA JSON:
{{"pertemuan": [{{"nomor": 1, "materi": "str", "durasi": "str", "kegiatan": [{{"fase": "PEMBUKAAN", "aktivitas": ["str", "str"], "waktu": "5'", "dl": "Meaningful"}}, {{"fase": "INTI (MEMAHAMI)", "aktivitas": ["str", "str"], "waktu": "15'", "dl": "Mindful"}}, {{"fase": "INTI (MENGAPLIKASIKAN)", "aktivitas": ["str", "str"], "waktu": "10'", "dl": "Joyful"}}, {{"fase": "PENUTUP", "aktivitas": ["str"], "waktu": "5'", "dl": "Mindful"}}]}}]}}"""

def prompt_step_3(form, step2):
    jumlah = form.get('jumlah_pertemuan', 1) 
    return f"""Tahap akhir modul {form['mapel']} bab {form['bab']}. Buat asesmen, LKPD (BUAT TEPAT {jumlah} LKPD, SATU UNTUK SETIAP PERTEMUAN), remedial/pengayaan, glosarium, daftar pustaka.
PENTING: Balas HANYA JSON VALID. Jangan gunakan kutip ganda (") di dalam nilai teks. "materi_ajar" cukup 1 paragraf padat agar tidak terpotong.
Balas HANYA JSON:
{{"penilaian": {{"awal": ["str"], "formatif": ["str"], "sumatif": ["str"]}}, "asesmen_lampiran": {{"awal_lisan": ["str"], "sumatif_hots": ["str"]}}, "materi_ajar": "str 1 paragraf padat", "lkpd": [{{"nomor": 1, "judul": "str", "memahami": "str", "mengaplikasikan": "str", "merefleksikan": "str"}}], "tindak_lanjut": {{"remedial": "str", "pengayaan": "str", "refleksi_siswa": ["str"], "refleksi_guru": ["str"]}}, "glosarium": [{{"istilah": "str", "definisi": "str"}}], "daftar_pustaka": ["str"]}}"""

# ==============================================================================
# PROMPT DOKUMEN LAIN
# ==============================================================================
def prompt_cp(form, d1_context=None, d2_context=None):
    return f"""Buat dokumen Capaian Pembelajaran (CP) Mapel {form['mapel']} Fase/Kelas {form['kelas']}. 
Sesuaikan dengan struktur berikut: A. Rasional, B. Tujuan, C. Tabel Elemen, D. Capaian Pembelajaran Fase.
Balas HANYA JSON: 
{{"rasional": "str 1 paragraf", "tujuan": ["str"], "elemen": [{{"nama": "str", "deskripsi": "str"}}], "cp_paragraf": "str 1 paragraf", "cp_tabel": [{{"elemen": "str", "capaian": "str"}}]}}"""

def prompt_atp(form, d1_context=None, d2_context=None):
    sinkron = get_sinkronisasi_context(d1_context, d2_context)
    return f"""Buat isi Alur Tujuan Pembelajaran (ATP) Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}.{sinkron}
Tabel ATP harus berisi No, Elemen, Tujuan Pembelajaran (TP) per Bab, Alur Tujuan Pembelajaran (ATP), Materi Pokok, dan Alokasi Waktu.
Balas HANYA JSON: 
{{"cp_fase": "str", "rows": [{{"no": "1", "elemen": "str", "tp": "str", "atp": "str", "materi": "str", "jp": "str"}}]}}"""

def prompt_prota(form, d1_context=None, d2_context=None):
    sinkron = get_sinkronisasi_context(d1_context, d2_context)
    return f"""Buat isi Program Tahunan (PROTA) Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}.{sinkron} Total JP harus mencakup semua materi. 
Kolom terdiri dari Semester (Ganjil/Genap), No, Materi Pokok/Bab, Alokasi Waktu (JP), dan Keterangan (contoh: "2 Pertemuan").
Balas HANYA JSON: {{"rows": [{{"semester": "1 (Ganjil)", "no": "1", "materi": "str", "jp": "str", "keterangan": "str"}}]}}"""

def prompt_promes(form, d1_context=None, d2_context=None):
    sinkron = get_sinkronisasi_context(d1_context, d2_context)
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    return f"""Buat rincian Program Semester Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Pecah ke bulan {bulan}. "minggu" array angka minggu (1-5).{sinkron} 
Balas HANYA JSON: {{"rows": [{{"materi_tp": "str", "jp": "str", "bulan": "Juli", "minggu": [1, 2]}}]}}"""

def prompt_kktp(form, d1_context=None, d2_context=None):
    sinkron = get_sinkronisasi_context(d1_context, d2_context)
    return f"""Buat Kriteria Ketercapaian Tujuan Pembelajaran (KKTP) Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}.{sinkron}
Tabel berisi 2 kolom: Tujuan Pembelajaran dan Kriteria Ketercapaian (Indikator). 
Balas HANYA JSON: {{"rows": [{{"tp": "str", "kriteria": "str"}}]}}"""

def prompt_jurnal(form, d1_context=None, d2_context=None):
    sinkron = get_sinkronisasi_context(d1_context, d2_context)
    return f"""Buat isi Jurnal Mengajar Harian untuk TEPAT {form['jumlah_pertemuan']} pertemuan. Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}.{sinkron} Balas HANYA JSON: {{"rows": [{{"pertemuan": "1", "topik": "str", "aktivitas": "str (Aktivitas Deep Learning)", "asesmen": "str"}}]}}"""

def prompt_lkpd(form, d1_context=None, d2_context=None):
    sinkron = get_sinkronisasi_context(d1_context, d2_context)
    return f"""Buat Lembar Kerja Peserta Didik (LKPD) lengkap yang interaktif untuk {form['jumlah_pertemuan']} pertemuan. Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}.{sinkron}
LKPD ini akan dicetak dan dibagikan ke siswa. Balas HANYA JSON:
{{"lkpd": [{{"pertemuan": 1, "topik": "str", "tujuan_kegiatan": "str", "alat_bahan": ["str"], "langkah_kerja": ["str"], "soal_latihan": ["str"]}}]}}"""


# ==============================================================================
# FUNGSI PEMBANTU FORMATTING DOCX
# ==============================================================================
def safe_list(val, default=None):
    if default is None: default = ["-"]
    if val is None: return default
    if isinstance(val, str): return [val]
    if isinstance(val, list) and len(val) > 0: return val
    return default

def set_cell_background(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)

def style_cell(cell, text, bold=False, color="000000", center=False, size=10, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = p.add_run(str(text))
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)

def doc_bullet_style(cell):
    return cell.paragraphs[0].style

def banner(doc, text, hex_color, size=12):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    table.columns[0].width = Cm(18)
    cell = table.rows[0].cells[0]
    set_cell_background(cell, hex_color)
    style_cell(cell, text, bold=True, color="FFFFFF", size=size)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table

def field_table(doc):
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(4.5)
    table.columns[1].width = Cm(13.5)
    return table

def add_field_row(table, label, content_items):
    row = table.add_row()
    label_cell, value_cell = row.cells[0], row.cells[1]
    label_cell.width = Cm(4.5); value_cell.width = Cm(13.5)
    set_cell_background(label_cell, COLOR_LABEL)
    set_cell_background(value_cell, COLOR_VALUE)
    style_cell(label_cell, label, bold=True, size=10.5)

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

def create_base_doc(landscape=False):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(1.5); section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5); section.bottom_margin = Cm(1.5)
    if landscape:
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width
    return doc

def add_signatures(doc, form, full_width=False):
    doc.add_paragraph("\n")
    sig_table = doc.add_table(rows=1, cols=2)
    w = Cm(12) if full_width else Cm(9)
    sig_table.columns[0].width = w; sig_table.columns[1].width = w
    p1 = sig_table.rows[0].cells[0].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run(f"Mengetahui,\nKepala Sekolah {form['sekolah']}\n\n\n\n")
    p1.add_run(f"({form['kepala_madrasah']})").bold = True
    p1.add_run("\nNIP. .....................................")
    
    p2 = sig_table.rows[0].cells[1].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"{form['titimangsa']}\nGuru Mata Pelajaran\n\n\n\n")
    p2.add_run(f"({form['penyusun']})").bold = True
    p2.add_run("\nNIP. .....................................")

def create_header(doc, title, form):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(14)
    doc.add_paragraph(f"Mata Pelajaran: {form['mapel']}")
    doc.add_paragraph(f"Nama Sekolah: {form['sekolah']}")
    if "Penyusun" in title or "ATP" in title or "KKTP" in title:
        doc.add_paragraph(f"Nama Penyusun: {form['penyusun']}")
    doc.add_paragraph(f"Fase/Kelas: {form['kelas']}")
    doc.add_paragraph(f"Tahun Ajaran: {form['tahun_pelajaran']}")
    doc.add_paragraph()

# ==============================================================================
# BUILDERS (Semua Builder Tersedia Sesuai Format Sebelumnya)
# ==============================================================================
def build_cover(form: dict, jenis_cover: str) -> bytes:
    is_landscape = (jenis_cover in ["Cover Program Tahunan & Semester", "Cover CP", "Cover ATP"])
    doc = create_base_doc(landscape=is_landscape)
    for _ in range(4): doc.add_paragraph()
    judul_utama = "BUKU PERANGKAT PEMBELAJARAN\n"
    if jenis_cover == "Cover Modul Ajar":
        judul_utama = "MODUL AJAR\nKURIKULUM BERBASIS CINTA\n"
    elif jenis_cover == "Cover Program Tahunan & Semester":
        judul_utama = "PROGRAM TAHUNAN DAN SEMESTER\n"
    elif jenis_cover == "Cover Jurnal Mengajar":
        judul_utama = "JURNAL MENGAJAR HARIAN\n"
    elif jenis_cover == "Cover CP":
        judul_utama = "CAPAIAN PEMBELAJARAN (CP)\n"
    elif jenis_cover == "Cover ATP":
        judul_utama = "ALUR TUJUAN PEMBELAJARAN (ATP)\n"
    elif jenis_cover == "Cover KKTP":
        judul_utama = "KRITERIA KETERCAPAIAN TUJUAN PEMBELAJARAN (KKTP)\n"
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(judul_utama)
    r1.bold = True
    r1.font.size = Pt(22)
    r1_sub = p1.add_run("(Pendekatan Deep Learning - KMA 1503/2025)")
    r1_sub.font.size = Pt(14)
    for _ in range(3): doc.add_paragraph()
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"Mata Pelajaran : {form['mapel']}\n")
    r2.bold = True
    r2.font.size = Pt(16)
    r2_sub = p2.add_run(f"Kelas / Fase : {form['kelas']}\nSemester : {form['semester']}")
    r2_sub.font.size = Pt(14)
    for _ in range(5): doc.add_paragraph()
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3_a = p3.add_run("Disusun Oleh:\n")
    r3_a.font.size = Pt(14)
    r3_b = p3.add_run(f"{form['penyusun']}")
    r3_b.bold = True
    r3_b.font.size = Pt(16)
    for _ in range(6): doc.add_paragraph()
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run(f"{form['sekolah']}\n")
    r4.bold = True
    r4.font.size = Pt(18)
    r4_sub = p4.add_run(f"Tahun Pelajaran {form['tahun_pelajaran']}")
    r4_sub.bold = True
    r4_sub.font.size = Pt(14)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_modul_ajar(form: dict, full_data: dict) -> bytes:
    doc = create_base_doc(landscape=False)
    d1 = full_data.get("step1", {})
    d2 = full_data.get("step2", {})
    d3 = full_data.get("step3", {})

    title_table = doc.add_table(rows=1, cols=1)
    cell = title_table.rows[0].cells[0]
    set_cell_background(cell, COLOR_TITLE)
    p1 = cell.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run("MODUL AJAR").bold = True
    p1.runs[0].font.size, p1.runs[0].font.color.rgb = Pt(16), RGBColor.from_string("FFFFFF")
    doc.add_paragraph()

    banner(doc, "IDENTITAS MODUL AJAR", COLOR_IDENTITY_HEAD)
    identity = field_table(doc)
    pertemuan_list = d2.get("pertemuan", [])
    if not isinstance(pertemuan_list, list): pertemuan_list = []
    
    for label, value in [
        ("Mata Pelajaran", form["mapel"]), ("Kelas / Fase", form["kelas"]),
        ("Semester", form["semester"]), ("Alokasi Waktu", f"{len(pertemuan_list)} Pertemuan x {form['alokasi']}"),
        ("Bab / Topik", form["bab"]), ("Penyusun", form["penyusun"]),
        ("Sekolah", form["sekolah"]), ("Tahun Pelajaran", form["tahun_pelajaran"])
    ]:
        add_field_row(identity, label, value)
    doc.add_paragraph()

    banner(doc, "A. IDENTIFIKASI PESERTA DIDIK", COLOR_SECTION_A)
    ident = field_table(doc)
    id_data = d1.get("identifikasi", {})
    if not isinstance(id_data, dict): id_data = {}
    add_field_row(ident, "Pengetahuan Awal", safe_list(id_data.get("pengetahuan_awal")))
    add_field_row(ident, "Minat Belajar", safe_list(id_data.get("minat_belajar")))
    add_field_row(ident, "Latar Belakang", str(id_data.get("latar_belakang", "-")))
    add_field_row(ident, "Kebutuhan Belajar", safe_list(id_data.get("kebutuhan_belajar")))
    add_field_row(ident, "Dimensi Profil Kelulusan", safe_list(id_data.get("dimensi_profil")))
    add_field_row(ident, "Topik Panca Cinta", safe_list(id_data.get("panca_cinta")))
    doc.add_paragraph()

    banner(doc, "B. DESAIN PEMBELAJARAN", COLOR_SECTION_B)
    desain = field_table(doc)
    ds_data = d1.get("desain", {})
    if not isinstance(ds_data, dict): ds_data = {}
    add_field_row(desain, "Capaian Pembelajaran (CP)", str(ds_data.get("capaian_pembelajaran", "-")))
    add_field_row(desain, "Tujuan Pembelajaran (TP)", safe_list(ds_data.get("tujuan_pembelajaran")))
    add_field_row(desain, "Lintas Disiplin Ilmu", safe_list(ds_data.get("lintas_disiplin")))
    add_field_row(desain, "Topik Pembelajaran", safe_list(ds_data.get("topik_pembelajaran")))
    add_field_row(desain, "Praktik Pedagogi", safe_list(ds_data.get("praktik_pedagogi")))
    add_field_row(desain, "Lingkungan Belajar", safe_list(ds_data.get("lingkungan_belajar")))
    add_field_row(desain, "Kemitraan Pembelajaran", safe_list(ds_data.get("kemitraan_pembelajaran")))
    add_field_row(desain, "Pemanfaatan Digital", safe_list(ds_data.get("pemanfaatan_digital")))
    doc.add_paragraph()

    for p in pertemuan_list:
        if not isinstance(p, dict): continue
        materi = p.get("materi", "Materi Pembelajaran")
        durasi = p.get("durasi", form['alokasi'])
        doc.add_heading(f"PENGALAMAN BELAJAR – PERTEMUAN {p.get('nomor', '1')}", level=2)
        doc.add_paragraph(f"Materi: {materi}\nDurasi: {durasi}")
        
        t_pb = doc.add_table(rows=1, cols=4)
        t_pb.style = 'Table Grid'
        t_pb.columns[0].width = Cm(3.5); t_pb.columns[1].width = Cm(10.0)
        t_pb.columns[2].width = Cm(1.5); t_pb.columns[3].width = Cm(3.0)
        
        hdr = t_pb.rows[0].cells
        headers = ["FASE KEGIATAN", "AKTIVITAS PEMBELAJARAN", "WAKTU", "PRINSIP DL"]
        for i in range(4):
            set_cell_background(hdr[i], COLOR_LABEL)
            style_cell(hdr[i], headers[i], bold=True, center=True)
            
        kegiatan_list = p.get("kegiatan", [])
        if isinstance(kegiatan_list, list):
            for keg in kegiatan_list:
                if not isinstance(keg, dict): continue
                row = t_pb.add_row()
                row.cells[0].text = str(keg.get("fase", ""))
                
                akt_list = keg.get("aktivitas", [])
                if isinstance(akt_list, list):
                    txt_akt = "\n".join([f"- {a}" for a in akt_list])
                else:
                    txt_akt = str(akt_list)
                
                row.cells[1].text = txt_akt
                row.cells[2].text = str(keg.get("waktu", "-"))
                row.cells[3].text = str(keg.get("dl", "-"))
        doc.add_paragraph()

    banner(doc, "PENILAIAN / ASESMEN", COLOR_IDENTITY_HEAD)
    t_penilaian = field_table(doc)
    pen = d3.get("penilaian", {})
    if not isinstance(pen, dict): pen = {}
    add_field_row(t_penilaian, "Asesmen Awal (Diagnostik)", safe_list(pen.get("awal")))
    add_field_row(t_penilaian, "Asesmen Formatif", safe_list(pen.get("formatif")))
    add_field_row(t_penilaian, "Asesmen Sumatif", safe_list(pen.get("sumatif")))
    doc.add_paragraph()

    banner(doc, "LAMPIRAN I – ASESMEN", COLOR_LAMPIRAN_I)
    asesmen_lamp = d3.get("asesmen_lampiran", {})
    if not isinstance(asesmen_lamp, dict): asesmen_lamp = {}
    
    doc.add_heading("A. ASESMEN AWAL (LISAN)", level=3)
    for a in safe_list(asesmen_lamp.get("awal_lisan")): doc.add_paragraph(f"• {a}")
    
    doc.add_heading("B. RUBRIK PENILAIAN SIKAP (Skala 1-4)", level=3)
    t_sikap = doc.add_table(rows=5, cols=5)
    t_sikap.style = 'Table Grid'
    h_sikap = ["Aspek Sikap", "Skor 4 (Sangat Baik)", "Skor 3 (Baik)", "Skor 2 (Cukup)", "Skor 1 (Perlu Bimb.)"]
    for i, h in enumerate(h_sikap):
        set_cell_background(t_sikap.cell(0, i), COLOR_LABEL)
        style_cell(t_sikap.cell(0, i), h, bold=True, center=True)
    sikap_data = [
        ["Disiplin", "Selalu hadir & taat", "Hadir tepat waktu", "Sering terlambat", "Sering absen"],
        ["Tanggung Jawab", "Tugas tepat & baik", "Tugas selesai", "Sering terlambat", "Tidak dikerjakan"],
        ["Kerjasama", "Sangat aktif", "Aktif", "Kurang aktif", "Tidak peduli"],
        ["Toleransi", "Sangat menghargai", "Menghargai", "Kurang menghargai", "Tidak menghargai"]
    ]
    for r_idx, row_data in enumerate(sikap_data, start=1):
        for c_idx, cell_data in enumerate(row_data):
            style_cell(t_sikap.cell(r_idx, c_idx), cell_data)
            
    doc.add_heading("C. ASESMEN SUMATIF (SOAL HOTS)", level=3)
    for i, a in enumerate(safe_list(asesmen_lamp.get("sumatif_hots")), 1): doc.add_paragraph(f"{i}. {a}")
    doc.add_paragraph()

    banner(doc, "LAMPIRAN II – MATERI AJAR", COLOR_LAMPIRAN_II)
    doc.add_paragraph(str(d3.get("materi_ajar", "-")))
    doc.add_paragraph()
    
    banner(doc, "LAMPIRAN III – LKPD (LEMBAR KERJA PESERTA DIDIK)", COLOR_LAMPIRAN_III)
    doc.add_paragraph("(Catatan: Untuk format siap cetak, silakan gunakan file dokumen LKPD Cetak yang ter-generate secara terpisah)")
    lkpd_data = d3.get("lkpd", [])
    if isinstance(lkpd_data, list):
        for p in lkpd_data:
            if not isinstance(p, dict): continue
            doc.add_heading(f"LKPD PERTEMUAN {p.get('nomor', '')} – {p.get('judul', 'Tugas')}", level=3)
            doc.add_paragraph("Pedoman: Memahami (40) + Mengaplikasikan (40) + Merefleksikan (20) = 100")
            
            t_lkpd = doc.add_table(rows=3, cols=2)
            t_lkpd.style = 'Table Grid'
            t_lkpd.columns[0].width = Cm(4.5); t_lkpd.columns[1].width = Cm(13.5)
            
            for i, (k, v) in enumerate([("MEMAHAMI", p.get("memahami", "")), 
                                        ("MENGAPLIKASIKAN", p.get("mengaplikasikan", "")), 
                                        ("MEREFLEKSIKAN", p.get("merefleksikan", ""))]):
                set_cell_background(t_lkpd.cell(i, 0), COLOR_LABEL)
                style_cell(t_lkpd.cell(i, 0), k, bold=True)
                style_cell(t_lkpd.cell(i, 1), str(v))
            doc.add_paragraph()

    banner(doc, "LAMPIRAN V – TINDAK LANJUT DAN REFLEKSI", COLOR_LAMPIRAN_V)
    tl = d3.get("tindak_lanjut", {})
    if not isinstance(tl, dict): tl = {}
    doc.add_heading("A. PROGRAM REMEDIAL", level=3); doc.add_paragraph(str(tl.get("remedial", "-")))
    doc.add_heading("B. PROGRAM PENGAYAAN", level=3); doc.add_paragraph(str(tl.get("pengayaan", "-")))
    doc.add_heading("C. REFLEKSI", level=3)
    doc.add_paragraph("Refleksi Peserta Didik:")
    for r in safe_list(tl.get("refleksi_siswa")): doc.add_paragraph(f"- {r}")
    doc.add_paragraph("Refleksi Guru:")
    for r in safe_list(tl.get("refleksi_guru")): doc.add_paragraph(f"- {r}")
    doc.add_paragraph()

    banner(doc, "GLOSARIUM & DAFTAR PUSTAKA", COLOR_TITLE)
    doc.add_heading("GLOSARIUM", level=3)
    glosarium_data = d3.get("glosarium", [])
    if isinstance(glosarium_data, list):
        for g in glosarium_data:
            if isinstance(g, dict):
                doc.add_paragraph(f"• {g.get('istilah', '')}: {str(g.get('definisi', ''))}")
            
    doc.add_heading("DAFTAR PUSTAKA", level=3)
    for dp in safe_list(d3.get("daftar_pustaka")): doc.add_paragraph(f"- {dp}")

    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_cp(form, ai_data):
    doc = create_base_doc(landscape=False)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("CAPAIAN PEMBELAJARAN (CP)")
    run.bold = True
    run.font.size = Pt(14)
    doc.add_paragraph(f"Mata Pelajaran: {form['mapel']}")
    doc.add_paragraph(f"Fase / Kelas: {form['kelas']}")
    doc.add_paragraph()
    doc.add_heading("A. Rasional Mata Pelajaran", level=3)
    doc.add_paragraph(ai_data.get("rasional", ""))
    doc.add_heading("B. Tujuan Mata Pelajaran", level=3)
    doc.add_paragraph(f"Mata pelajaran {form['mapel']} bertujuan agar peserta didik dapat:")
    for t in safe_list(ai_data.get("tujuan", [])):
        p = doc.add_paragraph()
        p.style = 'List Bullet'
        p.add_run(t)
    doc.add_heading("C. Elemen-elemen Mata Pelajaran", level=3)
    t_elemen = doc.add_table(rows=1, cols=2)
    t_elemen.style = 'Table Grid'
    set_cell_background(t_elemen.cell(0, 0), "EFEFEF"); style_cell(t_elemen.cell(0, 0), "Elemen", bold=True)
    set_cell_background(t_elemen.cell(0, 1), "EFEFEF"); style_cell(t_elemen.cell(0, 1), "Deskripsi", bold=True)
    t_elemen.columns[0].width = Cm(4.0); t_elemen.columns[1].width = Cm(14.0)
    for el in safe_list(ai_data.get("elemen", [])):
        if isinstance(el, dict):
            r = t_elemen.add_row().cells
            style_cell(r[0], el.get("nama", "")); style_cell(r[1], el.get("deskripsi", ""))
    doc.add_paragraph()
    doc.add_heading(f"D. Capaian Pembelajaran Fase", level=3)
    doc.add_paragraph(ai_data.get("cp_paragraf", ""))
    t_cp = doc.add_table(rows=1, cols=2)
    t_cp.style = 'Table Grid'
    set_cell_background(t_cp.cell(0, 0), "EFEFEF"); style_cell(t_cp.cell(0, 0), "Elemen", bold=True)
    set_cell_background(t_cp.cell(0, 1), "EFEFEF"); style_cell(t_cp.cell(0, 1), "Capaian Pembelajaran", bold=True)
    t_cp.columns[0].width = Cm(4.0); t_cp.columns[1].width = Cm(14.0)
    for cp in safe_list(ai_data.get("cp_tabel", [])):
        if isinstance(cp, dict):
            r = t_cp.add_row().cells
            style_cell(r[0], cp.get("elemen", "")); style_cell(r[1], cp.get("capaian", ""))
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_atp(form, ai_data):
    doc = create_base_doc(landscape=True)
    create_header(doc, "ALUR TUJUAN PEMBELAJARAN (ATP)", form)
    cp_fase = ai_data.get("cp_fase", "")
    if cp_fase:
        p = doc.add_paragraph()
        p.add_run("Capaian Pembelajaran Fase:\n").bold = True
        p.add_run(cp_fase)
        doc.add_paragraph()
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    headers = ["No.", "Elemen", "Tujuan Pembelajaran (TP) per Bab", "Alur Tujuan Pembelajaran (ATP)", "Materi Pokok", "Alokasi Waktu (JP)"]
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), "EFEFEF")
        style_cell(table.cell(0, i), h, bold=True, center=True)
    table.columns[0].width = Cm(1.0); table.columns[5].width = Cm(2.0)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("no", ""), center=True)
        style_cell(r[1], row.get("elemen", ""))
        style_cell(r[2], row.get("tp", ""))
        style_cell(r[3], row.get("atp", ""))
        style_cell(r[4], row.get("materi", ""))
        style_cell(r[5], row.get("jp", ""), center=True)
    add_signatures(doc, form, full_width=True)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_prota(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "PROGRAM TAHUNAN (PROTA)", form)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Semester", "No", "Materi Pokok / Bab", "Alokasi Waktu (JP)", "Keterangan"]
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), "EFEFEF")
        style_cell(table.cell(0, i), h, bold=True, center=True)
    table.columns[0].width = Cm(2.5); table.columns[1].width = Cm(1.0); table.columns[3].width = Cm(3.0)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("semester", ""), center=True) 
        style_cell(r[1], row.get("no", ""), center=True)
        style_cell(r[2], row.get("materi", "")) 
        style_cell(r[3], row.get("jp", ""), center=True)
        style_cell(r[4], row.get("keterangan", ""), center=True)
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_promes(form, ai_data):
    doc = create_base_doc(landscape=True)
    create_header(doc, "PROGRAM SEMESTER (PROSEM)", form)
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    total_cols = 2 + (len(bulan) * 5)
    table = doc.add_table(rows=2, cols=total_cols)
    table.style = 'Table Grid'
    table.cell(0, 0).merge(table.cell(1, 0)); style_cell(table.cell(0, 0), "Materi / Tujuan Pembelajaran", bold=True, center=True)
    table.cell(0, 1).merge(table.cell(1, 1)); style_cell(table.cell(0, 1), "JP", bold=True, center=True)
    set_cell_background(table.cell(0, 0), "EFEFEF"); set_cell_background(table.cell(0, 1), "EFEFEF")
    table.columns[0].width = Cm(6.0); table.columns[1].width = Cm(1.5)
    col_idx = 2
    for b in bulan:
        table.cell(0, col_idx).merge(table.cell(0, col_idx + 4))
        style_cell(table.cell(0, col_idx), b, bold=True, center=True)
        set_cell_background(table.cell(0, col_idx), "EFEFEF")
        for w in range(5):
            style_cell(table.cell(1, col_idx + w), str(w + 1), bold=True, center=True)
            set_cell_background(table.cell(1, col_idx + w), "F5F5F5")
            table.columns[col_idx + w].width = Cm(0.6)
        col_idx += 5
    row_m = table.add_row().cells
    style_cell(row_m[0], "Minggu ke-", bold=True, italic=True)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("materi_tp", ""))
        style_cell(r[1], row.get("jp", ""), center=True)
        target_bulan = row.get("bulan", "")
        minggu_aktif = row.get("minggu", [])
        if not isinstance(minggu_aktif, list): minggu_aktif = []
        idx = 2
        for b in bulan:
            for w in range(1, 6):
                if target_bulan.lower() == b.lower() and w in minggu_aktif:
                    set_cell_background(r[idx], COLOR_TITLE) 
                idx += 1
    add_signatures(doc, form, full_width=True)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_kktp(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "KRITERIA KETERCAPAIAN TUJUAN PEMBELAJARAN (KKTP)", form)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    headers = ["Tujuan Pembelajaran (TP)", "Kriteria Ketercapaian (Indikator)"]
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), "EFEFEF")
        style_cell(table.cell(0, i), h, bold=True, center=True)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("tp", ""))
        style_cell(r[1], row.get("kriteria", ""))
    doc.add_paragraph("\nKeterangan Tingkat Ketercapaian:").bold = True
    doc.add_paragraph("Perlu Bimbingan: Peserta didik belum mampu memenuhi kriteria dan memerlukan bimbingan pada hampir seluruh bagian.")
    doc.add_paragraph("Cukup: Peserta didik mampu memenuhi sebagian kriteria namun belum konsisten atau masih memerlukan sedikit bantuan.")
    doc.add_paragraph("Baik: Peserta didik mampu memenuhi seluruh kriteria yang ditetapkan secara mandiri.")
    doc.add_paragraph("Sangat Baik: Peserta didik mampu memenuhi seluruh kriteria dengan analisis yang lebih mendalam, kritis, dan mampu mengaplikasikannya pada konteks baru.")
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_jurnal(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "JURNAL MENGAJAR HARIAN", form)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Pertemuan", "Topik / Materi", "Aktivitas Deep Learning", "Asesmen", "Ket/Paraf"]
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), COLOR_TITLE)
        style_cell(table.cell(0, i), h, bold=True, color="FFFFFF", center=True)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("pertemuan", ""), center=True)
        style_cell(r[1], row.get("topik", "")); style_cell(r[2], row.get("aktivitas", ""))
        style_cell(r[3], row.get("asesmen", "")); style_cell(r[4], "") 
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_lkpd(form, ai_data):
    doc = create_base_doc(landscape=False)
    lkpd_list = ai_data.get("lkpd", [])
    if not isinstance(lkpd_list, list): lkpd_list = []
    for item in lkpd_list:
        if not isinstance(item, dict): continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("LEMBAR KERJA PESERTA DIDIK (LKPD)\n")
        r.bold = True
        r.font.size = Pt(16)
        p.add_run(f"Mata Pelajaran: {form['mapel']} | Kelas: {form['kelas']}")
        doc.add_heading(f"LKPD Pertemuan {item.get('pertemuan', '')} - {item.get('topik', '')}", level=2)
        table = doc.add_table(rows=3, cols=1)
        table.style = 'Table Grid'
        p1 = table.cell(0, 0).paragraphs[0]
        p1.add_run("Nama Kelompok / Siswa  : ................................................................................").bold = True
        p2 = table.cell(1, 0).paragraphs[0]
        p2.add_run("Kelas                  : ................................................................................").bold = True
        p3 = table.cell(2, 0).paragraphs[0]
        p3.add_run("Hari, Tanggal          : ................................................................................").bold = True
        doc.add_paragraph()
        doc.add_heading("A. Tujuan Kegiatan", level=3)
        doc.add_paragraph(str(item.get("tujuan_kegiatan", "-")))
        doc.add_heading("B. Alat dan Bahan (Jika Ada)", level=3)
        for ab in safe_list(item.get("alat_bahan")): doc.add_paragraph(f"- {ab}")
        doc.add_heading("C. Langkah Kerja", level=3)
        for i, lk in enumerate(safe_list(item.get("langkah_kerja")), 1): doc.add_paragraph(f"{i}. {lk}")
        doc.add_heading("D. Tugas / Soal Latihan", level=3)
        for i, soal in enumerate(safe_list(item.get("soal_latihan")), 1):
            doc.add_paragraph(f"{i}. {soal}")
            for _ in range(4): doc.add_paragraph()
        doc.add_page_break()
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()


# ==============================================================================
# UI STREAMLIT (ALUR EKSEKUSI BERTAHAP / ANTI-ERROR)
# ==============================================================================
st.title("📘 MI MIFTAHUSSALAM ADMIN GURU GENERATOR V.3 ")
st.markdown("*Berbasis Model Lagos AI 9.1 - Alur Bertahap Anti-Limit*")

with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Pecahan dan Desimal")
        kelas = st.selectbox("Jenjang / Kelas", list(JENJANG_FASE.keys()), index=6) 
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan (Modul, Jurnal, LKPD)", min_value=1, max_value=8, value=2)
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
        penyusun = st.text_input("Penyusun (Guru)", placeholder="Erian Kurniawan, S.E")
    with col4:
        kepala_madrasah = st.text_input("Kepala Madrasah", placeholder="Drs. Andi Supriadi")

    st.divider()
    
    jenis_cover = st.selectbox(
        "Pilih Jenis Cover yang Ingin Dibuat",
        [
            "Tanpa Cover", 
            "Cover Administrasi Guru (Buku Umum)",
            "Cover Modul Ajar",
            "Cover Program Tahunan & Semester",
            "Cover Jurnal Mengajar",
            "Cover CP",
            "Cover ATP",
            "Cover KKTP"
        ]
    )
    
    pilihan_dokumen = st.multiselect(
        "Pilih dokumen yang ingin di-generate otomatis",
        ["Modul Ajar", "Capaian Pembelajaran (CP)", "Alur Tujuan Pembelajaran (ATP)", "Prota", "Promes", "KKTP", "Jurnal Mengajar", "LKPD Siswa (Cetak)"],
        default=["Modul Ajar", "Capaian Pembelajaran (CP)"]
    )

    submitted = st.form_submit_button("✨ Eksekusi & Generate Dokumen", use_container_width=True)

if submitted:
    if not (mapel and bab and penyusun and sekolah and kepala_madrasah):
        st.warning("Lengkapi minimal Mata Pelajaran, Bab/Topik, Penyusun, Kepala, dan Sekolah.")
    elif not pilihan_dokumen and jenis_cover == "Tanpa Cover":
        st.warning("Silakan pilih minimal 1 dokumen atau cover yang ingin dibuat.")
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
            # === 1. GENERATE COVER ===
            if jenis_cover != "Tanpa Cover":
                cover_status = st.empty()
                cover_status.info(f"📄 Membuat {jenis_cover}...")
                doc_bytes = build_cover(form, jenis_cover)
                filename = f"{jenis_cover.replace(' & ', '_').replace(' ', '_')}_{safe_mapel}_{safe_kelas}.docx"
                st.session_state["hasil_generate"][filename] = doc_bytes
                cover_status.success(f"✅ Selesai membuat {jenis_cover}")
                time.sleep(1)

            # === 2. BIKIN "OTAK UTAMA" (TAHAP 1 & 2 DARI MODUL AJAR) ===
            st.markdown("### 🧠 Proses Pembuatan Master Data & Modul Ajar (3 Tahap)")
            
            # Tahap 1
            step1_status = st.empty()
            step1_status.info("⏳ Tahap 1/3: Merancang Desain Pembelajaran (Identitas, CP, TP, & Pilar KBC)...")
            d1_context = call_ai(prompt_step_1(form))
            step1_status.success("✅ Tahap 1 Selesai!")
            time.sleep(2) # Memberikan jeda (delay) agar API tidak mencapai batas rate-limit
            
            # Tahap 2
            step2_status = st.empty()
            step2_status.info("⏳ Tahap 2/3: Menyusun Pengalaman Belajar (Deep Learning) & Rincian Pertemuan...")
            d2_context = call_ai(prompt_step_2(form, d1_context))
            step2_status.success("✅ Tahap 2 Selesai!")
            time.sleep(2)

            # === 3. GENERATE DOKUMEN BERTAHAP (ANTREAN) ===
            st.markdown("### ⚙️ Pemrosesan Dokumen Antrean (Satu-per-Satu)")
            progress_bar = st.progress(0)
            total_docs = len(pilihan_dokumen)
            
            for i, tipe in enumerate(pilihan_dokumen):
                doc_status = st.empty()
                
                try:
                    if tipe == "Modul Ajar":
                        doc_status.info(f"🔄 Sedang memproses: {tipe} (Tahap 3/3: Asesmen, Penilaian, LKPD & Lampiran)...")
                        d3_context = call_ai(prompt_step_3(form, d2_context))
                        doc_bytes = build_modul_ajar(form, {"step1": d1_context, "step2": d2_context, "step3": d3_context})
                    elif tipe == "Capaian Pembelajaran (CP)":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_cp(form, d1_context, d2_context))
                        doc_bytes = build_cp(form, ai_data)
                    elif tipe == "Alur Tujuan Pembelajaran (ATP)":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_atp(form, d1_context, d2_context))
                        doc_bytes = build_atp(form, ai_data)
                    elif tipe == "Prota":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_prota(form, d1_context, d2_context))
                        doc_bytes = build_prota(form, ai_data)
                    elif tipe == "Promes":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_promes(form, d1_context, d2_context))
                        doc_bytes = build_promes(form, ai_data)
                    elif tipe == "KKTP":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_kktp(form, d1_context, d2_context))
                        doc_bytes = build_kktp(form, ai_data)
                    elif tipe == "Jurnal Mengajar":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_jurnal(form, d1_context, d2_context))
                        doc_bytes = build_jurnal(form, ai_data)
                    elif tipe == "LKPD Siswa (Cetak)":
                        doc_status.info(f"🔄 Sedang memproses: {tipe}...")
                        ai_data = call_ai(prompt_lkpd(form, d1_context, d2_context))
                        doc_bytes = build_lkpd(form, ai_data)
                    
                    # Menyimpan hasil ke session state
                    safe_tipe = tipe.replace(" & ", "_").replace(" ", "_").replace("(", "").replace(")", "")
                    filename = f"{safe_tipe}_{safe_mapel}_{safe_kelas}.docx"
                    st.session_state["hasil_generate"][filename] = doc_bytes
                    
                    doc_status.success(f"✅ Berhasil membuat dokumen: {tipe}")
                    time.sleep(3) # Jeda ekstra 3 detik setiap kali selesai membuat 1 dokumen (mencegah error API).
                    
                except Exception as e:
                    doc_status.error(f"❌ Gagal memproses {tipe}: {e}")
                    
                progress_bar.progress((i + 1) / total_docs)
            
            st.success("🎉 Selesai! Semua dokumen berhasil digenerate secara bertahap tanpa error.")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan sistem: {e}")

if "hasil_generate" in st.session_state and st.session_state["hasil_generate"]:
    st.divider()
    st.write("⬇️ **Silakan Unduh File Anda di Bawah Ini:**")
    for fname, fbytes in st.session_state["hasil_generate"].items():
        st.download_button(label=f"Unduh {fname}", data=fbytes, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
