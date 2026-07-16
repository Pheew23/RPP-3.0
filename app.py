"""
Generator Modul Ajar berbasis AI - Versi Prompt Chaining
-------------------------------------------------------
Streamlit app memecah pemanggilan AI menjadi 3 tahap agar hasil lebih
detail, tidak terpotong (error JSON), dan sesuai format aslinya.
"""

import io
import json
import re
import time

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Palet warna
COLOR_TITLE = "1F4E79"
COLOR_IDENTITY_HEAD = "2E75B6"
COLOR_LABEL = "DEEAF1"
COLOR_VALUE = "FFFFFF"
COLOR_SECTION = "1F4E79"
COLOR_MEETING = "843C0C"

MODEL_NAME = "nvidia/nemotron-3-ultra-550b-a55b"
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"

FASE_MAP = {1: "A", 2: "A", 3: "B", 4: "B", 5: "C", 6: "C"}

st.set_page_config(page_title="Generator Modul Ajar AI", page_icon="📘", layout="wide")

@st.cache_resource
def get_client():
    api_key = st.secrets.get("NVIDIA_API_KEY")
    if not api_key:
        st.error("NVIDIA_API_KEY belum ada di st.secrets.")
        st.stop()
    return OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)

def call_ai(prompt: str, temperature=0.2) -> dict:
    """Fungsi pemanggilan AI yang aman dari error Markdown/JSON"""
    client = get_client()
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        temperature=temperature,
        max_tokens=6000,
    )
    text = response.choices[0].message.content.strip()
    
    st.session_state["raw_ai_output"] = text # Untuk debugging
    
    text = re.sub(r'^```json\s*', '', text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r'^```\s*', '', text, flags=re.MULTILINE)
    text = text.strip()
    
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group(0)
        
    return json.loads(text)

# ==============================================================================
# PROMPT CHAINING SCHEMAS
# ==============================================================================

def prompt_step_1(form: dict) -> str:
    return f"""Kamu pakar kurikulum SD Indonesia. Buat Bagian A (Identifikasi) dan C (Desain Pembelajaran) 
untuk Mapel: {form['mapel']}, Kelas: {form['kelas']}, Topik: {form['bab']}.

Tuliskan detail yang mendalam, sesuaikan dengan Kurikulum Berbasis Cinta.
Balas HANYA dengan JSON valid sesuai skema:
{{
  "identifikasi": {{
    "pengetahuan_awal": ["string", "string"],
    "minat_belajar": ["string", "string"],
    "latar_belakang": "string",
    "kebutuhan_belajar": ["string", "string"],
    "dimensi_profil": ["string", "string"],
    "panca_cinta": ["string", "string"]
  }},
  "desain": {{
    "capaian": "string",
    "tujuan": ["string", "string"],
    "lintas_disiplin": ["string", "string"],
    "topik_pembelajaran": ["string", "string"]
  }}
}}"""

def prompt_step_2(form: dict, step1_data: dict) -> str:
    n = form["jumlah_pertemuan"]
    return f"""Melanjutkan modul {form['mapel']} Kelas {form['kelas']} bab {form['bab']}.
Tujuan: {step1_data['desain']['tujuan'][0]}. 
Buat detail Pengalaman Belajar untuk TEPAT {n} pertemuan. 
Wajib integrasikan sintak Problem Based Learning (PBL) dan prinsip Deep Learning (Mindful, Meaningful, Joyful).

Balas HANYA dengan JSON valid sesuai skema berikut:
{{
  "pertemuan": [
    {{
      "nomor": 1,
      "materi": "string",
      "pembukaan": [{{"kegiatan": "string", "dl": "Meaningful / Joyful / Mindful"}}],
      "inti": [{{"sintak": "Orientasi / Organisasi / Penyelidikan / Penyajian", "kegiatan": "string", "dl": "string"}}],
      "penutup": [{{"kegiatan": "string", "dl": "string"}}]
    }}
  ]
}}"""

def prompt_step_3(form: dict, step2_data: dict) -> str:
    return f"""Tahap akhir modul {form['mapel']} bab {form['bab']}. 
Buat Lampiran (Asesmen, LKPD, Tindak Lanjut). LKPD harus merujuk pada kegiatan di pertemuan.

Balas HANYA dengan JSON valid sesuai skema:
{{
  "asesmen": {{
    "awal_lisan": ["string", "string"],
    "sumatif_hots": ["string", "string"]
  }},
  "lkpd": [
    {{
      "pertemuan": 1,
      "judul": "string",
      "tugas_memahami": "string",
      "tugas_mengaplikasikan": "string",
      "tugas_merefleksikan": "string"
    }}
  ],
  "tindak_lanjut": {{
    "remedial": "string",
    "pengayaan": "string"
  }}
}}"""

# ==============================================================================
# DOCX BUILDER (DI-UPGRADE SESUAI CONTOH)
# ==============================================================================

def set_cell_background(cell, hex_color: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)

def build_docx(form: dict, full_data: dict) -> bytes:
    doc = Document()
    doc.sections[0].left_margin = Cm(2)
    doc.sections[0].right_margin = Cm(2)
    
    d1 = full_data["step1"]
    d2 = full_data["step2"]
    d3 = full_data["step3"]

    # --- JUDUL ---
    doc.add_heading('MODUL AJAR KURIKULUM BERBASIS CINTA', level=1)
    doc.add_heading('PENDEKATAN DEEP LEARNING', level=2)
    doc.add_paragraph()

    # --- IDENTITAS ---
    doc.add_heading('IDENTITAS MODUL AJAR', level=2)
    table = doc.add_table(rows=5, cols=4)
    table.style = 'Table Grid'
    records = [
        ("Mata Pelajaran", form['mapel'], "Kelas / Fase", f"Kelas {form['kelas']}"),
        ("Semester", form['semester'], "Topik / Bab", form['bab']),
        ("Model", form['model'], "Metode", form['metode']),
        ("Penyusun", form['penyusun'], "Sekolah", form['sekolah']),
        ("Tahun Pelajaran", form['tahun_pelajaran'], "Alokasi", form['alokasi'])
    ]
    for i, row in enumerate(records):
        for j, text in enumerate(row):
            table.rows[i].cells[j].text = str(text)
    doc.add_paragraph()

    # --- BAGIAN A & C ---
    doc.add_heading('A. IDENTIFIKASI PESERTA DIDIK', level=2)
    for key, vals in d1['identifikasi'].items():
        doc.add_heading(key.replace('_', ' ').title(), level=3)
        if isinstance(vals, list):
            for v in vals: doc.add_paragraph(f"• {v}")
        else:
            doc.add_paragraph(str(vals))
            
    doc.add_heading('C. DESAIN PEMBELAJARAN', level=2)
    for key, vals in d1['desain'].items():
        doc.add_heading(key.replace('_', ' ').title(), level=3)
        if isinstance(vals, list):
            for v in vals: doc.add_paragraph(f"• {v}")
        else:
            doc.add_paragraph(str(vals))

    # --- PENGALAMAN BELAJAR (PBL & DL) ---
    doc.add_page_break()
    doc.add_heading('PENGALAMAN BELAJAR', level=1)
    
    for p in d2['pertemuan']:
        doc.add_heading(f"PERTEMUAN {p['nomor']}: {p.get('materi', 'Materi')}", level=2)
        
        # Tabel Sintak PBL
        t_pbl = doc.add_table(rows=1, cols=3)
        t_pbl.style = 'Table Grid'
        hdr = t_pbl.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = "SINTAK / FASE", "AKTIVITAS PEMBELAJARAN", "PRINSIP DL"
        set_cell_background(hdr[0], COLOR_LABEL)
        set_cell_background(hdr[1], COLOR_LABEL)
        set_cell_background(hdr[2], COLOR_LABEL)

        # Pembukaan
        row = t_pbl.add_row()
        row.cells[0].text = "PEMBUKAAN"
        row.cells[1].text = "\n".join([f"- {k['kegiatan']}" for k in p['pembukaan']])
        row.cells[2].text = p['pembukaan'][0].get('dl', 'Mindful')

        # Inti
        for i, inti in enumerate(p['inti']):
            row = t_pbl.add_row()
            row.cells[0].text = f"INTI:\n{inti.get('sintak', 'Penyelidikan')}"
            row.cells[1].text = inti.get('kegiatan', '')
            row.cells[2].text = inti.get('dl', 'Meaningful')

        # Penutup
        row = t_pbl.add_row()
        row.cells[0].text = "PENUTUP"
        row.cells[1].text = "\n".join([f"- {k['kegiatan']}" for k in p['penutup']])
        row.cells[2].text = p['penutup'][0].get('dl', 'Joyful')
        doc.add_paragraph()

    # --- LAMPIRAN ---
    doc.add_page_break()
    doc.add_heading('LAMPIRAN I - ASESMEN', level=2)
    doc.add_heading('Asesmen Awal Lisan', level=3)
    for a in d3['asesmen']['awal_lisan']: doc.add_paragraph(f"• {a}")
    doc.add_heading('Asesmen Sumatif (HOTS)', level=3)
    for a in d3['asesmen']['sumatif_hots']: doc.add_paragraph(f"• {a}")

    doc.add_heading('LAMPIRAN II - LKPD', level=2)
    for l in d3['lkpd']:
        doc.add_heading(f"LKPD Pertemuan {l['pertemuan']}: {l.get('judul', '')}", level=3)
        doc.add_paragraph(f"1. Memahami:\n{l.get('tugas_memahami', '')}")
        doc.add_paragraph(f"2. Mengaplikasikan:\n{l.get('tugas_mengaplikasikan', '')}")
        doc.add_paragraph(f"3. Merefleksikan:\n{l.get('tugas_merefleksikan', '')}")

    doc.add_heading('LAMPIRAN III - TINDAK LANJUT', level=2)
    doc.add_paragraph(f"Remedial:\n{d3['tindak_lanjut'].get('remedial', '')}")
    doc.add_paragraph(f"Pengayaan:\n{d3['tindak_lanjut'].get('pengayaan', '')}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ==============================================================================
# UI STREAMLIT
# ==============================================================================
st.title("📘 Generator Modul Ajar (3-Step Prompt Chaining)")

with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Bahasa Indonesia")
        bab = st.text_input("Bab / Topik", placeholder="Bab 1: Anak-Anak yang Mengubah Dunia")
        kelas = st.selectbox("Kelas", [1, 2, 3, 4, 5, 6])
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
        jumlah_pertemuan = st.number_input("Jumlah Pertemuan", min_value=1, max_value=6, value=4)
    with col2:
        alokasi = st.text_input("Alokasi Waktu per Pertemuan", value="4 JP x 35 menit")
        model_pembelajaran = st.text_input("Model Pembelajaran", value="PBL (Problem Based Learning)")
        metode = st.text_input("Metode Pembelajaran", value="Ceramah Interaktif, Diskusi Kelompok")
        penyusun = st.text_input("Penyusun (Nama Guru)")
        sekolah = st.text_input("Sekolah")
        tahun_pelajaran = st.text_input("Tahun Pelajaran", value="2026/2027")

    submitted = st.form_submit_button("✨ Generate Full Modul (Butuh ~1 Menit)", use_container_width=True)

if submitted:
    if not (mapel and bab and penyusun):
        st.warning("Lengkapi data yang kosong!")
    else:
        form = dict(
            mapel=mapel, bab=bab, kelas=kelas, semester=semester,
            jumlah_pertemuan=int(jumlah_pertemuan), alokasi=alokasi,
            model=model_pembelajaran, metode=metode, penyusun=penyusun,
            sekolah=sekolah, tahun_pelajaran=tahun_pelajaran,
        )
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # TAHAP 1
            status_text.write("⏳ Langkah 1/3: Menyusun Identitas & Desain Pembelajaran...")
            d1 = call_ai(prompt_step_1(form))
            progress_bar.progress(33)
            
            # TAHAP 2
            status_text.write("⏳ Langkah 2/3: Merancang Pengalaman Belajar & PBL...")
            d2 = call_ai(prompt_step_2(form, d1))
            progress_bar.progress(66)
            
            # TAHAP 3
            status_text.write("⏳ Langkah 3/3: Menyiapkan Asesmen & LKPD...")
            d3 = call_ai(prompt_step_3(form, d2))
            progress_bar.progress(100)
            
            status_text.success("✅ Modul Ajar Berhasil Disusun!")
            
            full_data = {"step1": d1, "step2": d2, "step3": d3}
            st.session_state["full_data"] = full_data
            st.session_state["form"] = form
            
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
    st.subheader(f"🎉 Selesai! Modul Ajar {form['mapel']} Siap Diunduh")
    
    docx_bytes = build_docx(form, full_data)
    st.download_button(
        "⬇️ Unduh sebagai Word (.docx)",
        data=docx_bytes,
        file_name=f"Modul_Ajar_{form['mapel']}_Bab_{form['bab'].replace(' ', '')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
