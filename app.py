
"""
Generator Dokumen Admin Guru MI (KBC & KMA 1503/2025)
--------------------------------------------------------------------------------
Pembaruan: Modul Ajar Super Detail, Fix JSON Parser, Fitur Dropdown Cover Dinamis
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

MODEL_NAME = "thinkingmachines/inkling"
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

def call_ai(prompt: str, temperature=0.3) -> dict:
    client = get_client()
    response = client.chat.completions.create(
        model=MODEL_NAME, messages=[{"role": "user", "content": prompt}],
        temperature=temperature, max_tokens=8192,
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
# PROMPT SCHEMAS (SUPER DETAIL)
# ==============================================================================
def prompt_step_1(form):
    return f"""Kamu pakar Kurikulum Merdeka Deep Learning Berbasis Cinta dengan 5 pilar (KBC). Buat Bagian A & B modul gunakan bahasa yang humanis agar tidak terlihat AI
untuk Mapel: {form['mapel']}, Jenjang: {form['kelas']}, Topik: {form['bab']}. PENTING: CP WAJIB sebutkan "KMA Nomor 1503 Tahun 2025".
PENTING: Balas HANYA dengan JSON valid. DILARANG menggunakan tanda kutip ganda (") di dalam teks string.
Balas HANYA JSON:
{{"identifikasi": {{"pengetahuan_awal": ["str"], "minat_belajar": ["str"], "latar_belakang": "str", "kebutuhan_belajar": ["str"], "dimensi_profil": ["str"], "panca_cinta": ["str"]}}, "desain": {{"capaian_pembelajaran": "str", "tujuan_pembelajaran": ["str"], "lintas_disiplin": ["str"], "topik_pembelajaran": ["str"], "praktik_pedagogi": ["str"], "lingkungan_belajar": ["str"], "kemitraan_pembelajaran": ["str"], "pemanfaatan_digital": ["str"]}}}}"""

def prompt_step_2(form, step1):
    return f"""Lanjutkan modul {form['mapel']} {form['kelas']} bab {form['bab']}. Buat Pengalaman Belajar untuk TEPAT {form['jumlah_pertemuan']} pertemuan. Format 4 elemen untuk setiap kegiatan: fase, aktivitas, waktu, dl.
PENTING: Balas HANYA dengan JSON valid. DILARANG menggunakan tanda kutip ganda (") di dalam teks string.
Balas HANYA JSON:
{{"pertemuan": [{{"nomor": 1, "materi": "str", "durasi": "str", "kegiatan": [{{"fase": "PEMBUKAAN", "aktivitas": ["str", "str"], "waktu": "5'", "dl": "Meaningful"}}, {{"fase": "INTI (MEMAHAMI)", "aktivitas": ["str", "str"], "waktu": "15'", "dl": "Mindful"}}, {{"fase": "INTI (MENGAPLIKASIKAN)", "aktivitas": ["str", "str"], "waktu": "10'", "dl": "Joyful"}}, {{"fase": "PENUTUP", "aktivitas": ["str"], "waktu": "5'", "dl": "Mindful"}}]}}]}}"""

def prompt_step_3(form, step2):
    # Menambahkan instruksi eksplisit jumlah LKPD yang harus dibuat
    jumlah = form.get('jumlah_pertemuan', 1) 
    
    return f"""Tahap akhir modul {form['mapel']} bab {form['bab']}. Buat asesmen, LKPD (BUAT TEPAT {jumlah} LKPD, SATU UNTUK SETIAP PERTEMUAN), remedial/pengayaan, glosarium, daftar pustaka.
PENTING: Balas HANYA JSON VALID. Jangan gunakan kutip ganda (") di dalam nilai teks. "materi_ajar" cukup 1 paragraf padat agar tidak terpotong.
Balas HANYA JSON:
{{"penilaian": {{"awal": ["str"], "formatif": ["str"], "sumatif": ["str"]}}, "asesmen_lampiran": {{"awal_lisan": ["str"], "sumatif_hots": ["str"]}}, "materi_ajar": "str 1 paragraf padat", "lkpd": [{{"nomor": 1, "judul": "str", "memahami": "str", "mengaplikasikan": "str", "merefleksikan": "str"}}], "tindak_lanjut": {{"remedial": "str", "pengayaan": "str", "refleksi_siswa": ["str"], "refleksi_guru": ["str"]}}, "glosarium": [{{"istilah": "str", "definisi": "str"}}], "daftar_pustaka": ["str"]}}"""
def prompt_cpatp(form):
    return f"""Buat isi CP berdasarkan KMA No.1503 tahun 2025 dan ATP Mapel {form['mapel']} {form['kelas']} Topik {form['bab']} berdasar KMA 1503/2025. Balas HANYA JSON: {{"rows": [{{"elemen": "str", "cp": "str", "tp": "str", "atp": "str", "jp": "str"}}]}}"""

def prompt_prota(form):
    return f"""Buat isi Program Tahunan Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Balas HANYA JSON: {{"rows": [{{"no": "1", "elemen_cp": "str", "topik_tp": "str", "jp": "str", "semester": "1"}}]}}"""

def prompt_promes(form):
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    return f"""Buat rincian Program Semester Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Pecah ke bulan {bulan}. "minggu" array angka minggu (1-5). Balas HANYA JSON: {{"rows": [{{"no": "1", "materi_tp": "str", "jp": "str", "bulan": "Juli", "minggu": [1, 2]}}]}}"""

def prompt_jurnal(form):
    return f"""Buat isi Jurnal Mengajar Harian {form['jumlah_pertemuan']} pertemuan. Mapel {form['mapel']} {form['kelas']} Topik {form['bab']}. Balas HANYA JSON: {{"rows": [{{"pertemuan": "1", "topik": "str", "aktivitas": "str (Deep Learning)", "asesmen": "str"}}]}}"""

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

def style_cell(cell, text, bold=False, color="000000", center=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = p.add_run(str(text))
    run.bold = bold
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

def add_signatures(doc, form):
    doc.add_paragraph("\n")
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.columns[0].width = Cm(9); sig_table.columns[1].width = Cm(9)
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
# BUILDER COVER (DINAMIS BERDASARKAN PILIHAN DROPDOWN)
# ==============================================================================
def build_cover(form: dict, jenis_cover: str) -> bytes:
    doc = create_base_doc(landscape=False)
    
    # Spasi atas
    for _ in range(4): doc.add_paragraph()
    
    # Tentukan Teks Judul Utama berdasarkan pilihan
    if jenis_cover == "Cover Modul Ajar":
        judul_utama = "MODUL AJAR\nKURIKULUM BERBASIS CINTA\n"
    elif jenis_cover == "Cover Program Tahunan & Semester":
        judul_utama = "PROGRAM TAHUNAN DAN SEMESTER\n"
    elif jenis_cover == "Cover Jurnal Mengajar":
        judul_utama = "JURNAL MENGAJAR HARIAN\n"
    elif jenis_cover == "Cover CP & ATP":
        judul_utama = "CAPAIAN PEMBELAJARAN (CP) & ALUR TUJUAN PEMBELAJARAN (ATP)\n"
    else:
        judul_utama = "BUKU PERANGKAT PEMBELAJARAN\nADMINISTRASI GURU\n"
        
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(judul_utama)
    r1.bold = True
    r1.font.size = Pt(22)
    
    r1_sub = p1.add_run("(Pendekatan Deep Learning - KMA 1503/2025)")
    r1_sub.font.size = Pt(14)
    
    for _ in range(3): doc.add_paragraph()
    
    # Identitas Mata Pelajaran
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"Mata Pelajaran : {form['mapel']}\n")
    r2.bold = True
    r2.font.size = Pt(16)
    
    r2_sub = p2.add_run(f"Kelas / Fase : {form['kelas']}\nSemester : {form['semester']}")
    r2_sub.font.size = Pt(14)
    
    for _ in range(5): doc.add_paragraph()
    
    # Identitas Penyusun
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3_a = p3.add_run("Disusun Oleh:\n")
    r3_a.font.size = Pt(14)
    
    r3_b = p3.add_run(f"{form['penyusun']}")
    r3_b.bold = True
    r3_b.font.size = Pt(16)
    
    for _ in range(6): doc.add_paragraph()
    
    # Identitas Sekolah
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

# ==============================================================================
# BUILDER MODUL AJAR 
# ==============================================================================
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
    p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run("KURIKULUM BERBASIS CINTA \u2013 PENDEKATAN DEEP LEARNING\nKementrian Agama").bold = True
    p2.runs[0].font.size, p2.runs[0].font.color.rgb = Pt(11), RGBColor.from_string("FFFFFF")
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
        doc.add_heading(f"PENGALAMAN BELAJAR \u2013 PERTEMUAN {p.get('nomor', '1')}", level=2)
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

    banner(doc, "LAMPIRAN I \u2013 ASESMEN", COLOR_LAMPIRAN_I)
    asesmen_lamp = d3.get("asesmen_lampiran", {})
    if not isinstance(asesmen_lamp, dict): asesmen_lamp = {}
    
    doc.add_heading("A. ASESMEN AWAL (LISAN)", level=3)
    for a in safe_list(asesmen_lamp.get("awal_lisan")): doc.add_paragraph(f"\u2022 {a}")
    
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

    banner(doc, "LAMPIRAN II \u2013 MATERI AJAR", COLOR_LAMPIRAN_II)
    doc.add_paragraph(str(d3.get("materi_ajar", "-")))
    doc.add_paragraph()
    
    banner(doc, "LAMPIRAN III \u2013 LKPD (LEMBAR KERJA PESERTA DIDIK)", COLOR_LAMPIRAN_III)
    lkpd_data = d3.get("lkpd", [])
    if isinstance(lkpd_data, list):
        for p in lkpd_data:
            if not isinstance(p, dict): continue
            doc.add_heading(f"LKPD PERTEMUAN {p.get('nomor', '')} \u2013 {p.get('judul', 'Tugas')}", level=3)
            doc.add_paragraph("Nama:\nKelas:\nTanggal:")
            doc.add_paragraph("Petunjuk: Kerjakan secara mandiri lalu diskusikan dengan kelompokmu.")
            
            t_lkpd = doc.add_table(rows=3, cols=2)
            t_lkpd.style = 'Table Grid'
            t_lkpd.columns[0].width = Cm(4.5); t_lkpd.columns[1].width = Cm(13.5)
            
            for i, (k, v) in enumerate([("MEMAHAMI", p.get("memahami", "")), 
                                        ("MENGAPLIKASIKAN", p.get("mengaplikasikan", "")), 
                                        ("MEREFLEKSIKAN", p.get("merefleksikan", ""))]):
                set_cell_background(t_lkpd.cell(i, 0), COLOR_LABEL)
                style_cell(t_lkpd.cell(i, 0), k, bold=True)
                style_cell(t_lkpd.cell(i, 1), str(v))
            doc.add_paragraph("Pedoman Penskoran: Memahami (40) + Mengaplikasikan (40) + Merefleksikan (20) = 100")
            doc.add_paragraph()

    banner(doc, "LAMPIRAN V \u2013 TINDAK LANJUT DAN REFLEKSI", COLOR_LAMPIRAN_V)
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
                doc.add_paragraph(f"\u2022 {g.get('istilah', '')}: {str(g.get('definisi', ''))}")
            
    doc.add_heading("DAFTAR PUSTAKA", level=3)
    for dp in safe_list(d3.get("daftar_pustaka")): doc.add_paragraph(f"- {dp}")

    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# ==============================================================================
# BUILDERS DOKUMEN LAIN (CP, PROTA, PROMES, JURNAL)
# ==============================================================================
def build_cpatp(form, ai_data):
    doc = create_base_doc(landscape=True)
    create_header(doc, "CAPAIAN PEMBELAJARAN (CP) & ALUR TUJUAN PEMBELAJARAN (ATP)", form)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Elemen", "Capaian Pembelajaran (CP)", "Tujuan Pembelajaran (TP)", "Alur Tujuan Pembelajaran (ATP)", "JP"]
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), COLOR_TITLE)
        style_cell(table.cell(0, i), h, bold=True, color="FFFFFF", center=True)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("elemen", "")); style_cell(r[1], row.get("cp", ""))
        style_cell(r[2], row.get("tp", "")); style_cell(r[3], row.get("atp", ""))
        style_cell(r[4], row.get("jp", ""), center=True)
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_prota(form, ai_data):
    doc = create_base_doc(landscape=False)
    create_header(doc, "PROGRAM TAHUNAN (PROTA)", form)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["No", "Elemen / CP", "Topik / Tujuan Pembelajaran", "JP", "Semester"]
    for i, h in enumerate(headers):
        set_cell_background(table.cell(0, i), COLOR_TITLE)
        style_cell(table.cell(0, i), h, bold=True, color="FFFFFF", center=True)
    table.columns[0].width = Cm(1.0); table.columns[3].width = Cm(1.5); table.columns[4].width = Cm(2.0)
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("no", ""), center=True); style_cell(r[1], row.get("elemen_cp", ""))
        style_cell(r[2], row.get("topik_tp", "")); style_cell(r[3], row.get("jp", ""), center=True)
        style_cell(r[4], row.get("semester", ""), center=True)
    add_signatures(doc, form)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

def build_promes(form, ai_data):
    doc = create_base_doc(landscape=True)
    create_header(doc, "PROGRAM SEMESTER (PROMES)", form)
    is_sem1 = "1" in form['semester']
    bulan = ["Juli", "Agustus", "September", "Oktober", "November", "Desember"] if is_sem1 else ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
    total_cols = 3 + (len(bulan) * 5) + 1 
    table = doc.add_table(rows=2, cols=total_cols)
    table.style = 'Table Grid'
    
    table.cell(0, 0).merge(table.cell(1, 0)); style_cell(table.cell(0, 0), "No", bold=True, center=True)
    table.cell(0, 1).merge(table.cell(1, 1)); style_cell(table.cell(0, 1), "Materi", bold=True, center=True)
    table.cell(0, 2).merge(table.cell(1, 2)); style_cell(table.cell(0, 2), "JP", bold=True, center=True)
    for i in range(3): set_cell_background(table.cell(0, i), COLOR_LABEL)
    
    col_idx = 3
    for b in bulan:
        table.cell(0, col_idx).merge(table.cell(0, col_idx + 4))
        style_cell(table.cell(0, col_idx), b, bold=True, center=True)
        set_cell_background(table.cell(0, col_idx), COLOR_LABEL)
        for w in range(5):
            style_cell(table.cell(1, col_idx + w), str(w + 1), bold=True, center=True)
            set_cell_background(table.cell(1, col_idx + w), "EFEFEF")
        col_idx += 5
        
    table.cell(0, total_cols - 1).merge(table.cell(1, total_cols - 1))
    style_cell(table.cell(0, total_cols - 1), "Ket.", bold=True, center=True)
    set_cell_background(table.cell(0, total_cols - 1), COLOR_LABEL)
    
    for row in safe_list(ai_data.get("rows"), []):
        if not isinstance(row, dict): continue
        r = table.add_row().cells
        style_cell(r[0], row.get("no", ""), center=True); style_cell(r[1], row.get("materi_tp", ""))
        style_cell(r[2], row.get("jp", ""), center=True)
        target_bulan = row.get("bulan", "")
        minggu_aktif = row.get("minggu", [])
        if not isinstance(minggu_aktif, list): minggu_aktif = []
        
        idx = 3
        for b in bulan:
            for w in range(1, 6):
                if target_bulan.lower() == b.lower() and w in minggu_aktif:
                    set_cell_background(r[idx], COLOR_TITLE) 
                idx += 1
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
st.title("📘 MI MIFTAHUSSALAM ADMIN GURU GENERATOR V.3 ")
st.markdown("*Berbasis Model Lagos AI 9.1*")

with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Matematika")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Pecahan dan Desimal")
        kelas = st.selectbox("Jenjang / Kelas", list(JENJANG_FASE.keys()), index=6) 
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan (Modul & Jurnal)", min_value=1, max_value=8, value=2)
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
    
    # === DROPDOWN COVER BARU YANG LEBIH SPESIFIK ===
    jenis_cover = st.selectbox(
        "Pilih Jenis Cover yang Ingin Dibuat",
        [
            "Tanpa Cover", 
            "Cover Administrasi Guru (Buku Umum)",
            "Cover Modul Ajar",
            "Cover Program Tahunan & Semester",
            "Cover Jurnal Mengajar",
            "Cover CP & ATP"
        ]
    )
    
    pilihan_dokumen = st.multiselect(
        "Pilih dokumen yang ingin di-generate otomatis (Pilih sesuai kebutuhan)",
        ["Modul Ajar", "CP & ATP", "Prota", "Promes", "Jurnal Mengajar"],
        default=["Modul Ajar"]
    )

    submitted = st.form_submit_button("✨ Eksekusi & Generate (Auto-Download)", use_container_width=True)

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
            # === GENERATE COVER BILA DIPILIH ===
            if jenis_cover != "Tanpa Cover":
                st.write(f"📄 **Bentar ya! Memproses {jenis_cover}...**")
                doc_bytes = build_cover(form, jenis_cover)
                
                safe_nama_cover = jenis_cover.replace(" & ", "_").replace(" ", "_")
                filename = f"{safe_nama_cover}_{safe_mapel}_{safe_kelas}.docx"
                
                st.session_state["hasil_generate"][filename] = doc_bytes
                trigger_download(doc_bytes, filename)
                time.sleep(1.5)
            
            # === GENERATE DOKUMEN LAINNYA ===
            for tipe in pilihan_dokumen:
                if tipe == "Modul Ajar":
                    st.write("📘 **Bentar ya! Memproses Modul Ajar...**")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.write("⏳ AI Sedang Memasak : Langkah 1/3: Sabar ya...")
                    d1 = call_ai(prompt_step_1(form))
                    progress_bar.progress(33)
                    
                    status_text.write("⏳ AI Sedang Memasak : Langkah 2/3: Sebentar lagi kok...")
                    d2 = call_ai(prompt_step_2(form, d1))
                    progress_bar.progress(66)
                    
                    status_text.write("⏳ AI Sedang Memasak : Langkah 3/3: Beneran ini terakhir, Tunggu ya...")
                    d3 = call_ai(prompt_step_3(form, d2))
                    progress_bar.progress(100)
                    status_text.success("✅ Modul Ajar selesai!")
                    
                    doc_bytes = build_modul_ajar(form, {"step1": d1, "step2": d2, "step3": d3})
                
                else:
                    st.write(f"📄 **Bentar ya! Memproses {tipe}...**")
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
            
            st.success("🎉 Selesai! Dokumen sudah terunduh (Nikmati Hasilnya Selagi Hangat!)")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses data: {e}")
            with st.expander("🔍 Lihat Hasil Mentah AI (Debugging)"):
                st.code(st.session_state.get("raw_ai_output", ""))

if "hasil_generate" in st.session_state and st.session_state["hasil_generate"]:
    st.divider()
    st.write("⬇️ **Tombol Unduh Manual (Jika pop-up auto-download diblokir browser):**")
    for fname, fbytes in st.session_state["hasil_generate"].items():
        st.download_button(label=f"Unduh {fname}", data=fbytes, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
