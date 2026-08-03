"""
Generator Dokumen Admin Guru MI (KBC & KMA 1503/2025)
--------------------------------------------------------------------------------
Pembaruan: Multi-Bab Generator, Sinkronisasi Konteks Global, Aturan Ketat KMA 1503.
"""

import io
import json
import re
import datetime
import time
import pandas as pd

import streamlit as st
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

st.set_page_config(page_title="MIFSAL ADMIN GURU V4 (MULTI-BAB)", page_icon="📘", layout="wide")

@st.cache_resource
def get_client():
    api_key = st.secrets.get("NVIDIA_API_KEY")
    if not api_key:
        st.error("NVIDIA_API_KEY belum ada di st.secrets.")
        st.stop()
    return OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)

def call_ai(prompt: str, temperature=0.7) -> dict:
    client = get_client()
    text = ""
    max_retries = 4
    
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=MODEL_NAME, messages=[{"role": "user", "content": prompt}],
                temperature=temperature, max_tokens=12192,
            )
            text = response.choices[0].message.content.strip()
            break
        except Exception as e:
            if "429" in str(e) or "Too Many Requests" in str(e):
                if attempt < max_retries - 1:
                    time.sleep(15)
                    continue
            raise e

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
# PROMPT SINKRONISASI TOTAL (MULTI-BAB CONTEXT)
# ==============================================================================
def get_aggregated_sinkronisasi_context(all_chapters_data):
    sinkron_text = ""
    total_jp = 0
    
    for idx, chap in enumerate(all_chapters_data):
        bab_name = chap["bab"]
        d1 = chap["d1"]
        d2 = chap["d2"]
        
        sinkron_text += f"\n--- INFORMASI BAB {idx+1}: {bab_name} ---\n"
        
        if d1 and "desain" in d1:
            cp = d1["desain"].get("capaian_pembelajaran", "")
            tp = d1["desain"].get("tujuan_pembelajaran", [])
            if isinstance(tp, list): tp = ", ".join(tp)
            sinkron_text += f"CP: '{cp}'\nTP: '{tp}'\n"
            
        if d2 and "pertemuan" in d2:
            materi_list = [f"Pertemuan {p.get('nomor', '')}: {p.get('materi', '')}" for p in d2.get("pertemuan", []) if isinstance(p, dict)]
            materi_str = " | ".join(materi_list)
            sinkron_text += f"Rincian Materi (Wajib dipakai berurutan): {materi_str}\n"

    if sinkron_text:
        return f"\n\n[SANGAT PENTING - MASTER SINKRONISASI DOKUMEN]\nKamu WAJIB mematuhi data dari seluruh Bab berikut agar semua dokumen dari Bab 1 hingga akhir selaras:\n{sinkron_text}\nJANGAN membuat materi atau TP di luar data di atas! Rangkum dan susun secara logis."
    return ""

# ==============================================================================
# PROMPT MODUL AJAR (PER BAB)
# ==============================================================================
def prompt_step_1(form, bab_name):
    return f"""Kamu pakar Kurikulum Merdeka Pendekatan Deep Learning Berbasis Cinta dengan 5 pilar (KBC). WAJIB JANGAN SAMPAI BUAT KESALAHAN.
Mapel: {form['mapel']}, Jenjang: {form['kelas']}, Topik Khusus Modul Ini: {bab_name}. 
PENTING: CP dan TP WAJIB mengacu pada "KMA Nomor 1503 Tahun 2025". Masukan minimal 3 Pemanfaatan Digital, serta Panca Cinta KBC.
Balas HANYA JSON:
{{"identifikasi": {{"pengetahuan_awal": ["str"], "minat_belajar": ["str"], "latar_belakang": "str", "kebutuhan_belajar": ["str"], "dimensi_profil": ["str"], "panca_cinta": ["str"]}}, "desain": {{"capaian_pembelajaran": "str", "tujuan_pembelajaran": ["str"], "lintas_disiplin": ["str"], "topik_pembelajaran": ["str"], "praktik_pedagogi": ["str"], "lingkungan_belajar": ["str"], "kemitraan_pembelajaran": ["str"], "pemanfaatan_digital": ["str"]}}}}"""

def prompt_step_2(form, bab_name, jumlah_pertemuan, step1):
    return f"""Lanjutkan modul {form['mapel']} {form['kelas']} bab {bab_name}. Buat Pengalaman Belajar untuk TEPAT {jumlah_pertemuan} pertemuan. Format 4 elemen untuk setiap kegiatan: fase, aktivitas, waktu, dl.
Balas HANYA JSON:
{{"pertemuan": [{{"nomor": 1, "materi": "str", "durasi": "str", "kegiatan": [{{"fase": "PEMBUKAAN", "aktivitas": ["str"], "waktu": "5'", "dl": "Meaningful"}}, {{"fase": "INTI (MEMAHAMI)", "aktivitas": ["str"], "waktu": "15'", "dl": "Mindful"}}, {{"fase": "INTI (MENGAPLIKASIKAN)", "aktivitas": ["str"], "waktu": "10'", "dl": "Joyful"}}, {{"fase": "PENUTUP", "aktivitas": ["str"], "waktu": "5'", "dl": "Mindful"}}]}}]}}"""

def prompt_step_3(form, bab_name, jumlah_pertemuan, step2):
    return f"""Tahap akhir modul {form['mapel']} bab {bab_name}. Buat asesmen, LKPD (BUAT TEPAT {jumlah_pertemuan} LKPD), remedial, glosarium. "materi_ajar" cukup 1 paragraf padat.
Balas HANYA JSON:
{{"penilaian": {{"awal": ["str"], "formatif": ["str"], "sumatif": ["str"]}}, "asesmen_lampiran": {{"awal_lisan": ["str"], "sumatif_hots": ["str"]}}, "materi_ajar": "str", "lkpd": [{{"nomor": 1, "judul": "str", "memahami": "str", "mengaplikasikan": "str", "merefleksikan": "str"}}], "tindak_lanjut": {{"remedial": "str", "pengayaan": "str", "refleksi_siswa": ["str"], "refleksi_guru": ["str"]}}, "glosarium": [{{"istilah": "str", "definisi": "str"}}], "daftar_pustaka": ["str"]}}"""

# ==============================================================================
# PROMPT DOKUMEN LAIN (TERINTEGRASI MULTI-BAB & KMA 1503)
# ==============================================================================
def prompt_cp(form, aggregated_context):
    return f"""Buat dokumen Capaian Pembelajaran (CP) Mapel {form['mapel']} Fase/Kelas {form['kelas']}. 
[ATURAN MUTLAK]: KONTEN CP WAJIB DAN HARUS SESUAI DENGAN KMA 1503 TAHUN 2025. {aggregated_context}
Balas HANYA JSON: 
{{"rasional": "str 1 paragraf", "tujuan": ["str"], "elemen": [{{"nama": "str", "deskripsi": "str"}}], "cp_paragraf": "str 1 paragraf", "cp_tabel": [{{"elemen": "str", "capaian": "str"}}]}}"""

def prompt_atp(form, aggregated_context):
    return f"""Buat isi Alur Tujuan Pembelajaran (ATP) Lengkap Mapel {form['mapel']} {form['kelas']}. 
[ATURAN MUTLAK]: STRUKTUR DAN KONTEN WAJIB MENGACU KMA 1503 TAHUN 2025. {aggregated_context}
Balas HANYA JSON: 
{{"cp_fase": "str", "rows": [{{"no": "1", "elemen": "str", "tp": "str", "atp": "str", "materi": "str", "jp": "str"}}]}}"""

def prompt_prota(form, aggregated_context):
    return f"""Buat Program Tahunan (PROTA) Mapel {form['mapel']} {form['kelas']}. Masukkan semua bab yang ada di master konteks.{aggregated_context}
Balas HANYA JSON: {{"rows": [{{"semester": "1", "no": "1", "materi": "str", "jp": "str", "keterangan": "str"}}]}}"""

def prompt_promes(form, aggregated_context):
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    return f"""Buat Program Semester Lengkap Mapel {form['mapel']} {form['kelas']}. Distribusikan materi dari semua bab ke bulan {bulan}.{aggregated_context} 
Balas HANYA JSON: {{"rows": [{{"materi_tp": "str", "jp": "str", "bulan": "Juli", "minggu": [1, 2]}}]}}"""

def prompt_kktp(form, aggregated_context):
    return f"""Buat KKTP Mapel {form['mapel']} {form['kelas']} untuk SEMUA BAB.{aggregated_context}
Balas HANYA JSON: {{"rows": [{{"tp": "str", "kriteria": "str"}}]}}"""

def prompt_jurnal(form, aggregated_context, total_pertemuan):
    return f"""Buat Jurnal Mengajar Harian gabungan untuk total {total_pertemuan} pertemuan dari seluruh bab.{aggregated_context} 
Balas HANYA JSON: {{"rows": [{{"pertemuan": "1", "topik": "str", "aktivitas": "str", "asesmen": "str"}}]}}"""

def prompt_lkpd_global(form, aggregated_context, total_pertemuan):
    return f"""Buat Buku LKPD Lengkap untuk total {total_pertemuan} pertemuan mencakup seluruh bab.{aggregated_context}
Balas HANYA JSON: {{"lkpd": [{{"pertemuan": 1, "topik": "str", "tujuan_kegiatan": "str", "alat_bahan": ["str"], "langkah_kerja": ["str"], "soal_latihan": ["str"]}}]}}"""


# ==============================================================================
# FUNGSI PEMBANTU FORMATTING DOCX (Tetap Sama)
# ==============================================================================
def safe_list(val, default=None):
    if default is None: default = ["-"]
    if val is None: return default
    if isinstance(val, str): return [val]
    if isinstance(val, list) and len(val) > 0: return val
    return default

def set_cell_background(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)

def style_cell(cell, text, bold=False, color="000000", center=False, size=10, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = p.add_run(str(text))
    run.bold = bold; run.italic = italic; run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)

def doc_bullet_style(cell): return cell.paragraphs[0].style

def banner(doc, text, hex_color, size=12):
    table = doc.add_table(rows=1, cols=1); table.autofit = True; table.columns[0].width = Cm(18)
    cell = table.rows[0].cells[0]
    set_cell_background(cell, hex_color)
    style_cell(cell, text, bold=True, color="FFFFFF", size=size)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table

def field_table(doc):
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(4.5); table.columns[1].width = Cm(13.5)
    return table

def add_field_row(table, label, content_items):
    row = table.add_row()
    label_cell, value_cell = row.cells[0], row.cells[1]
    label_cell.width = Cm(4.5); value_cell.width = Cm(13.5)
    set_cell_background(label_cell, COLOR_LABEL); set_cell_background(value_cell, COLOR_VALUE)
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
    p1 = sig_table.rows[0].cells[0].paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run(f"Mengetahui,\nKepala Sekolah {form['sekolah']}\n\n\n\n")
    p1.add_run(f"({form['kepala_madrasah']})").bold = True
    p1.add_run("\nNIP. .....................................")
    
    p2 = sig_table.rows[0].cells[1].paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"{form['titimangsa']}\nGuru Mata Pelajaran\n\n\n\n")
    p2.add_run(f"({form['penyusun']})").bold = True
    p2.add_run("\nNIP. .....................................")

def create_header(doc, title, form):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title); run.bold = True; run.font.size = Pt(14)
    doc.add_paragraph(f"Mata Pelajaran: {form['mapel']}")
    doc.add_paragraph(f"Nama Sekolah: {form['sekolah']}")
    if "Penyusun" in title or "ATP" in title or "KKTP" in title:
        doc.add_paragraph(f"Nama Penyusun: {form['penyusun']}")
    doc.add_paragraph(f"Fase/Kelas: {form['kelas']}")
    doc.add_paragraph(f"Tahun Ajaran: {form['tahun_pelajaran']}")
    doc.add_paragraph()

# ==============================================================================
# BUILDERS (Cover & Document Builders)
# ==============================================================================
def build_cover(form: dict, jenis_cover: str) -> bytes:
    is_landscape = (jenis_cover in ["Cover Program Tahunan & Semester", "Cover CP", "Cover ATP"])
    doc = create_base_doc(landscape=is_landscape)
    for _ in range(4): doc.add_paragraph()
    judul_utama = "BUKU PERANGKAT PEMBELAJARAN\n"
    if jenis_cover == "Cover Modul Ajar": judul_utama = "MODUL AJAR\nKURIKULUM BERBASIS CINTA\n"
    elif jenis_cover == "Cover Program Tahunan & Semester": judul_utama = "PROGRAM TAHUNAN DAN SEMESTER\n"
    elif jenis_cover == "Cover Jurnal Mengajar": judul_utama = "JURNAL MENGAJAR HARIAN\n"
    elif jenis_cover == "Cover CP": judul_utama = "CAPAIAN PEMBELAJARAN (CP)\n"
    elif jenis_cover == "Cover ATP": judul_utama = "ALUR TUJUAN PEMBELAJARAN (ATP)\n"
    elif jenis_cover == "Cover KKTP": judul_utama = "KRITERIA KETERCAPAIAN TUJUAN PEMBELAJARAN (KKTP)\n"
    
    p1 = doc.add_paragraph(); p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(judul_utama); r1.bold = True; r1.font.size = Pt(22)
    r1_sub = p1.add_run("(Pendekatan Deep Learning - KMA 1503/2025)"); r1_sub.font.size = Pt(14)
    for _ in range(3): doc.add_paragraph()
    
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"Mata Pelajaran : {form['mapel']}\n"); r2.bold = True; r2.font.size = Pt(16)
    r2_sub = p2.add_run(f"Kelas / Fase : {form['kelas']}\nSemester : {form['semester']}")
    r2_sub.font.size = Pt(14)
    for _ in range(5): doc.add_paragraph()
    
    p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3_a = p3.add_run("Disusun Oleh:\n"); r3_a.font.size = Pt(14)
    r3_b = p3.add_run(f"{form['penyusun']}"); r3_b.bold = True; r3_b.font.size = Pt(16)
    for _ in range(6): doc.add_paragraph()
    
    p4 = doc.add_paragraph(); p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run(f"{form['sekolah']}\n"); r4.bold = True; r4.font.size = Pt(18)
    r4_sub = p4.add_run(f"Tahun Pelajaran {form['tahun_pelajaran']}"); r4_sub.bold = True; r4_sub.font.size = Pt(14)
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_modul_ajar(form: dict, bab_name: str, jumlah_pertemuan: int, full_data: dict) -> bytes:
    doc = create_base_doc(landscape=False)
    d1 = full_data.get("d1", {})
    d2 = full_data.get("d2", {})
    d3 = full_data.get("d3", {})

    title_table = doc.add_table(rows=1, cols=1)
    cell = title_table.rows[0].cells[0]
    set_cell_background(cell, COLOR_TITLE)
    p1 = cell.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run("MODUL AJAR").bold = True
    p1.runs[0].font.size, p1.runs[0].font.color.rgb = Pt(16), RGBColor.from_string("FFFFFF")
    doc.add_paragraph()

    banner(doc, "IDENTITAS MODUL AJAR", COLOR_IDENTITY_HEAD)
    identity = field_table(doc)
    
    for label, value in [
        ("Mata Pelajaran", form["mapel"]), ("Kelas / Fase", form["kelas"]),
        ("Semester", form["semester"]), ("Alokasi Waktu", f"{jumlah_pertemuan} Pertemuan x {form['alokasi']}"),
        ("Bab / Topik", bab_name), ("Penyusun", form["penyusun"]),
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

    pertemuan_list = d2.get("pertemuan", [])
    if not isinstance(pertemuan_list, list): pertemuan_list = []
    
    for p in pertemuan_list:
        if not isinstance(p, dict): continue
        doc.add_heading(f"PENGALAMAN BELAJAR – PERTEMUAN {p.get('nomor', '1')}", level=2)
        doc.add_paragraph(f"Materi: {p.get('materi', 'Materi')}\nDurasi: {p.get('durasi', form['alokasi'])}")
        
        t_pb = doc.add_table(rows=1, cols=4)
        t_pb.style = 'Table Grid'
        t_pb.columns[0].width = Cm(3.5); t_pb.columns[1].width = Cm(10.0)
        t_pb.columns[2].width = Cm(1.5); t_pb.columns[3].width = Cm(3.0)
        
        hdr = t_pb.rows[0].cells
        for i, h in enumerate(["FASE KEGIATAN", "AKTIVITAS PEMBELAJARAN", "WAKTU", "PRINSIP DL"]):
            set_cell_background(hdr[i], COLOR_LABEL)
            style_cell(hdr[i], h, bold=True, center=True)
            
        kegiatan_list = p.get("kegiatan", [])
        if isinstance(kegiatan_list, list):
            for keg in kegiatan_list:
                if not isinstance(keg, dict): continue
                row = t_pb.add_row()
                row.cells[0].text = str(keg.get("fase", ""))
                
                akt_list = keg.get("aktivitas", [])
                txt_akt = "\n".join([f"- {a}" for a in akt_list]) if isinstance(akt_list, list) else str(akt_list)
                
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
    doc.add_heading("B. ASESMEN SUMATIF (SOAL HOTS)", level=3)
    for i, a in enumerate(safe_list(asesmen_lamp.get("sumatif_hots")), 1): doc.add_paragraph(f"{i}. {a}")
    doc.add_paragraph()

    banner(doc, "LAMPIRAN II – MATERI AJAR", COLOR_LAMPIRAN_II)
    doc.add_paragraph(str(d3.get("materi_ajar", "-")))
    doc.add_paragraph()
    
    banner(doc, "LAMPIRAN III – LKPD", COLOR_LAMPIRAN_III)
    lkpd_data = d3.get("lkpd", [])
    if isinstance(lkpd_data, list):
        for p in lkpd_data:
            if not isinstance(p, dict): continue
            doc.add_heading(f"LKPD PERTEMUAN {p.get('nomor', '')} – {p.get('judul', 'Tugas')}", level=3)
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

    banner(doc, "LAMPIRAN V – TINDAK LANJUT", COLOR_LAMPIRAN_V)
    tl = d3.get("tindak_lanjut", {})
    if not isinstance(tl, dict): tl = {}
    doc.add_heading("A. PROGRAM REMEDIAL", level=3); doc.add_paragraph(str(tl.get("remedial", "-")))
    doc.add_heading("B. PROGRAM PENGAYAAN", level=3); doc.add_paragraph(str(tl.get("pengayaan", "-")))
    doc.add_heading("C. REFLEKSI", level=3)
    for r in safe_list(tl.get("refleksi_siswa")): doc.add_paragraph(f"- {r}")
    doc.add_paragraph()

    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_cp(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "CAPAIAN PEMBELAJARAN (CP)", form)
    doc.add_heading("A. Rasional Mata Pelajaran", level=3)
    doc.add_paragraph(ai_data.get("rasional", ""))
    doc.add_heading("B. Tujuan Mata Pelajaran", level=3)
    for t in safe_list(ai_data.get("tujuan", [])):
        doc.add_paragraph(t, style='List Bullet')
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
    for i, h in enumerate(["Tujuan Pembelajaran (TP)", "Kriteria Ketercapaian (Indikator)"]):
        set_cell_background(table.cell(0, i), "EFEFEF")
        style_cell(table.cell(0, i), h, bold=True, center=True)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("tp", "")); style_cell(r[1], row.get("kriteria", ""))
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_jurnal(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "JURNAL MENGAJAR HARIAN (GABUNGAN)", form)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    for i, h in enumerate(["Pertemuan", "Topik / Materi", "Aktivitas", "Asesmen", "Ket/Paraf"]):
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
    for item in safe_list(ai_data.get("lkpd", [])):
        if not isinstance(item, dict): continue
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("LEMBAR KERJA PESERTA DIDIK (LKPD)\n"); r.bold = True; r.font.size = Pt(16)
        p.add_run(f"Mata Pelajaran: {form['mapel']} | Kelas: {form['kelas']}")
        doc.add_heading(f"LKPD Pertemuan {item.get('pertemuan', '')} - {item.get('topik', '')}", level=2)
        table = doc.add_table(rows=3, cols=1)
        table.style = 'Table Grid'
        table.cell(0, 0).paragraphs[0].add_run("Nama Kelompok / Siswa  : ...................................................").bold = True
        table.cell(1, 0).paragraphs[0].add_run("Kelas                  : ...................................................").bold = True
        table.cell(2, 0).paragraphs[0].add_run("Hari, Tanggal          : ...................................................").bold = True
        doc.add_paragraph()
        doc.add_heading("A. Tujuan Kegiatan", level=3); doc.add_paragraph(str(item.get("tujuan_kegiatan", "-")))
        doc.add_heading("B. Alat dan Bahan (Jika Ada)", level=3)
        for ab in safe_list(item.get("alat_bahan")): doc.add_paragraph(f"- {ab}")
        doc.add_heading("C. Langkah Kerja", level=3)
        for i, lk in enumerate(safe_list(item.get("langkah_kerja")), 1): doc.add_paragraph(f"{i}. {lk}")
        doc.add_heading("D. Tugas / Soal Latihan", level=3)
        for i, soal in enumerate(safe_list(item.get("soal_latihan")), 1):
            doc.add_paragraph(f"{i}. {soal}")
            for _ in range(3): doc.add_paragraph()
        doc.add_page_break()
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# ==============================================================================
# UI STREAMLIT (SISTEM MULTI-BAB)
# ==============================================================================
st.title("📘 MI MIFTAHUSSALAM ADMIN GURU GENERATOR V.4 ")
st.markdown("*Berbasis Model Lagos AI 9.1 - Multi-Bab & Terintegrasi Master KMA 1503*")

with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Contoh: Fikih")
        kelas = st.selectbox("Jenjang / Kelas", list(JENJANG_FASE.keys()), index=6) 
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        alokasi = st.text_input("Alokasi Waktu per Pertemuan", value="4 JP x 35 menit")
        sekolah = st.text_input("Sekolah", value="MI Miftahussalam")
        tahun_pelajaran = st.text_input("Tahun Pelajaran", value="2026/2027")

    st.divider()
    st.markdown("### 📚 Data Bab Pembelajaran (Multi-Bab)")
    st.info("Tambahkan baris baru di tabel ini untuk membuat lebih dari 1 Bab secara otomatis. Modul Ajar akan di-generate untuk masing-masing Bab, sedangkan Dokumen seperti ATP/Prota akan dirangkum dari seluruh Bab di bawah ini.")
    
    df_bab_default = pd.DataFrame([
        {"Bab / Topik": "Bab 1: Ketentuan Zakat Fitrah", "Jumlah Pertemuan": 2},
        {"Bab / Topik": "Bab 2: Ketentuan Infak dan Sedekah", "Jumlah Pertemuan": 2}
    ])
    
    tabel_bab = st.data_editor(df_bab_default, num_rows="dynamic", use_container_width=True)

    st.divider()
    col3, col4 = st.columns(2)
    now = datetime.datetime.now()
    bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    titimangsa_otomatis = f"Bogor, {now.day} {bulan_indo[now.month - 1]} {now.year}"

    with col3:
        titimangsa = st.text_input("Titimangsa", value=titimangsa_otomatis)
        penyusun = st.text_input("Penyusun (Guru)", placeholder="Nama Guru, S.Pd.")
    with col4:
        kepala_madrasah = st.text_input("Kepala Madrasah", placeholder="Nama Kepala Sekolah")

    st.divider()
    
    jenis_cover = st.selectbox(
        "Pilih Jenis Cover yang Ingin Dibuat",
        ["Tanpa Cover", "Cover Administrasi Guru (Buku Umum)", "Cover Modul Ajar", "Cover Program Tahunan & Semester", "Cover Jurnal Mengajar", "Cover CP", "Cover ATP", "Cover KKTP"]
    )
    
    pilihan_dokumen = st.multiselect(
        "Pilih dokumen yang ingin di-generate otomatis",
        ["Modul Ajar (Per Bab)", "Capaian Pembelajaran (CP)", "Alur Tujuan Pembelajaran (ATP)", "Prota", "Promes", "KKTP", "Jurnal Mengajar (Global)", "LKPD Siswa (Global)"],
        default=["Modul Ajar (Per Bab)", "Capaian Pembelajaran (CP)", "Alur Tujuan Pembelajaran (ATP)"]
    )

    submitted = st.form_submit_button("✨ Eksekusi & Generate Semua Dokumen", use_container_width=True)

if submitted:
    if not (mapel and penyusun and sekolah and kepala_madrasah):
        st.warning("Lengkapi data identitas (Mapel, Penyusun, Sekolah, Kepala).")
    elif tabel_bab.empty or tabel_bab["Bab / Topik"].isnull().all():
        st.warning("Harap isi setidaknya 1 Bab pada tabel.")
    else:
        form = dict(
            mapel=mapel, kelas=kelas, semester=semester,
            alokasi=alokasi, penyusun=penyusun, sekolah=sekolah, 
            tahun_pelajaran=tahun_pelajaran, titimangsa=titimangsa, kepala_madrasah=kepala_madrasah
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
                st.session_state["hasil_generate"][f"{jenis_cover}_{safe_mapel}.docx"] = doc_bytes
                cover_status.success(f"✅ Selesai membuat {jenis_cover}")

            # === 2. PROSES MODUL AJAR (LOOPING MULTI-BAB) ===
            all_chapters_data = []
            total_pertemuan_global = 0
            
            # Konversi dataframe pandas ke list of dict
            daftar_bab = tabel_bab.dropna(subset=["Bab / Topik"]).to_dict('records')
            
            if "Modul Ajar (Per Bab)" in pilihan_dokumen or True: # Tetap harus di proses untuk master konteks
                st.markdown("### 🧠 Proses Pembuatan Master Data & Modul Ajar (Per Bab)")
                
                for idx, row in enumerate(daftar_bab):
                    bab_name = str(row["Bab / Topik"])
                    try: 
                        jml_pert = int(row["Jumlah Pertemuan"])
                    except: 
                        jml_pert = 2
                    
                    total_pertemuan_global += jml_pert
                    safe_bab = re.sub(r'[^a-zA-Z0-9_\-]', '_', bab_name)
                    
                    st.write(f"🔄 **Sedang memproses {bab_name}...**")
                    
                    st.info("Tahap 1: Desain Pembelajaran...")
                    d1_ctx = call_ai(prompt_step_1(form, bab_name))
                    time.sleep(3)
                    
                    st.info("Tahap 2: Pengalaman Belajar & Pertemuan...")
                    d2_ctx = call_ai(prompt_step_2(form, bab_name, jml_pert, d1_ctx))
                    time.sleep(3)
                    
                    st.info("Tahap 3: Asesmen & Lampiran...")
                    d3_ctx = call_ai(prompt_step_3(form, bab_name, jml_pert, d2_ctx))
                    
                    chapter_data = {"bab": bab_name, "jml_pert": jml_pert, "d1": d1_ctx, "d2": d2_ctx, "d3": d3_ctx}
                    all_chapters_data.append(chapter_data)
                    
                    # Generate File Modul Ajar khusus bab ini jika dipilih
                    if "Modul Ajar (Per Bab)" in pilihan_dokumen:
                        doc_bytes = build_modul_ajar(form, bab_name, jml_pert, chapter_data)
                        filename = f"Modul_Ajar_{safe_mapel}_Bab_{idx+1}_{safe_bab}.docx"
                        st.session_state["hasil_generate"][filename] = doc_bytes
                        st.success(f"✅ Berhasil membuat Modul Ajar: {bab_name}")
                    
                    time.sleep(5) # Jeda antar bab
            
            # === 3. SINKRONISASI KONTEKS GLOBAL ===
            master_context = get_aggregated_sinkronisasi_context(all_chapters_data)

            # === 4. GENERATE DOKUMEN ADMINISTRASI LAIN (MEMAKAI MASTER CONTEXT) ===
            st.markdown("### ⚙️ Pemrosesan Dokumen Administrasi Sinkronisasi")
            dokumen_admin = [d for d in pilihan_dokumen if d != "Modul Ajar (Per Bab)"]
            
            for tipe in dokumen_admin:
                doc_status = st.empty()
                try:
                    doc_status.info(f"🔄 Mensinkronkan & memproses: {tipe} (Sesuai KMA 1503)...")
                    
                    if tipe == "Capaian Pembelajaran (CP)":
                        ai_data = call_ai(prompt_cp(form, master_context))
                        doc_bytes = build_cp(form, ai_data)
                    elif tipe == "Alur Tujuan Pembelajaran (ATP)":
                        ai_data = call_ai(prompt_atp(form, master_context))
                        doc_bytes = build_atp(form, ai_data)
                    elif tipe == "Prota":
                        ai_data = call_ai(prompt_prota(form, master_context))
                        doc_bytes = build_prota(form, ai_data)
                    elif tipe == "Promes":
                        ai_data = call_ai(prompt_promes(form, master_context))
                        doc_bytes = build_promes(form, ai_data)
                    elif tipe == "KKTP":
                        ai_data = call_ai(prompt_kktp(form, master_context))
                        doc_bytes = build_kktp(form, ai_data)
                    elif tipe == "Jurnal Mengajar (Global)":
                        ai_data = call_ai(prompt_jurnal(form, master_context, total_pertemuan_global))
                        doc_bytes = build_jurnal(form, ai_data)
                    elif tipe == "LKPD Siswa (Global)":
                        ai_data = call_ai(prompt_lkpd_global(form, master_context, total_pertemuan_global))
                        doc_bytes = build_lkpd(form, ai_data)
                    
                    safe_tipe = tipe.replace(" & ", "_").replace(" ", "_").replace("(", "").replace(")", "")
                    st.session_state["hasil_generate"][f"{safe_tipe}_{safe_mapel}_Full_Semester.docx"] = doc_bytes
                    
                    doc_status.success(f"✅ Berhasil membuat & mensinkronkan: {tipe}")
                    time.sleep(8) 
                    
                except Exception as e:
                    doc_status.error(f"❌ Gagal memproses {tipe}: {e}")
            
            st.success("🎉 Selesai! Seluruh Ekosistem Dokumen Administrasi Anda Berhasil Dibuat dan Disinkronkan.")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan sistem: {e}")

if "hasil_generate" in st.session_state and st.session_state["hasil_generate"]:
    st.divider()
    st.write("⬇️ **Silakan Unduh File Anda di Bawah Ini:**")
    for fname, fbytes in st.session_state["hasil_generate"].items():
        st.download_button(label=f"Unduh {fname}", data=fbytes, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
