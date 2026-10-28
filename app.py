import os
import io
import datetime
import requests
import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION, WD_ORIENTATION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from streamlit_cookies_controller import CookieController
import docx

# ----------------- SAYFA YAPILANDIRMASI -----------------
st.set_page_config(
    page_title="İSG Saha Gözetim Raporu",
    page_icon="🛡️",
    layout="centered"
)

st.markdown("""
<style>
    .block-container {
        max-width: 860px !important;
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        margin: 0 auto !important;
    }
    .preview-box {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    .preview-box img {
        width: 100px !important;
        height: 100px !important;
        object-fit: cover !important;
        border-radius: 6px !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# ----------------- ÇEREZ & SESSION YÖNETİMİ -----------------
cookie_controller = CookieController()

if "findings" not in st.session_state:
    st.session_state.findings = []
if "api_key" not in st.session_state:
    st.session_state.api_key = ""
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

saved_cookie_key = cookie_controller.get("gemini_api_key")
if saved_cookie_key and not st.session_state.api_key:
    st.session_state.api_key = saved_cookie_key

# ----------------- SOL MENÜ: API AYARLARI -----------------
with st.sidebar:
    st.header("⚙️ Sistem Ayarları")
    st.caption("API Anahtarınız bulut sunucusuna gitmez; yalnızca tarayıcınızın çerezlerinde saklanır.")
    
    input_key = st.text_input(
        "Gemini API Anahtarı",
        type="password",
        value=st.session_state.api_key,
        help="Google AI Studio anahtarınızı buraya girip kaydedin."
    )
    
    col_save, col_clear = st.columns(2)
    with col_save:
        if st.button("💾 Çereze Kaydet", use_container_width=True):
            if input_key.strip():
                st.session_state.api_key = input_key.strip()
                cookie_controller.set("gemini_api_key", input_key.strip(), max_age=30*24*60*60)
                st.success("Tarayıcıya kaydedildi!")
            else:
                st.warning("Lütfen bir anahtar girin.")
                
    with col_clear:
        if st.button("🗑️ Çerezi Sil", use_container_width=True):
            cookie_controller.remove("gemini_api_key")
            st.session_state.api_key = ""
            st.rerun()

    active_api_key = input_key.strip() or st.session_state.api_key.strip()
    if active_api_key:
        st.success("✅ API Anahtarı Etkin")
    else:
        st.error("⚠️ Lütfen API anahtarınızı giriniz.")

    st.markdown("---")
    st.info(f"Eklenen Madde Sayısı: **{len(st.session_state.findings)}**")

# ----------------- WORD YARDIMCI FONKSİYONLARI -----------------
def set_cell_background(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=140):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_table_borders(table, color="000000"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def add_page_number_to_footer(footer):
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    fld_page = OxmlElement('w:fldSimple')
    fld_page.set(qn('w:instr'), 'PAGE')
    p._p.append(fld_page)
    
    run_sep = p.add_run(" / ")
    run_sep.font.name = "Times New Roman"
    run_sep.font.size = Pt(10)
    
    fld_numpages = OxmlElement('w:fldSimple')
    fld_numpages.set(qn('w:instr'), 'NUMPAGES')
    p._p.append(fld_numpages)

def add_cover_line(doc, text, size=24, bold=False, space_after=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    r.bold = bold
    return p

# ----------------- GEMINI API İLE METİN DÜZENLEME -----------------
def call_gemini_ai(raw_note, api_key):
    system_prompt = (
        "Sen profesyonel bir İSG Saha Denetim Danışmanısın. "
        "Gelen gözlem notlarını profesyonel bir rapora dönüştür. "
        "Şu kurallara SIKI SIKIYA uy:\n"
        "1. TESPİT: Sadece fiziksel gözlemi içermeli. Risk tanımlarını yazma. "
        "TESPİT METNİNDEKİ SON CÜMLE MUTLAKA '... tespit edilmiştir.' İFADESİ İLE BİTMELİDİR.\n"
        "2. ÖNERİ: Danışmanlık dili kullan. 'Derhal', 'Acilen' gibi emir kipleri yerine "
        "'Yapılması önerilmektedir', 'Sağlanması uygun olacaktır' gibi ifadeler kullan.\n"
        "3. YASAL DAYANAK: Konuyla ilgili TÜM yasal dayanakları (Kanun ve Yönetmelikler) bul. "
        "Sadece isimlerini yaz. Birden fazla varsa bunları her biri ayrı satırda olacak şekilde alt alta listele.\n"
        "Yanıtı SADECE şu JSON formatında ver: {\"tespit\": \"...\", \"oneri\": \"...\", \"yasal\": \"Yönetmelik 1\\nYönetmelik 2\"}"
    )

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": [{"text": f"Gözlem: \"{raw_note}\""}]}],
        "systemInstruction": {"parts": [{"text": system_prompt}]},
        "generationConfig": {"responseMimeType": "application/json"}
    }
    
    response = requests.post(url, json=payload, timeout=30)
    if response.status_code != 200:
        raise Exception(f"Google API Hatası: {response.text}")
    
    import json
    return json.loads(response.json()["candidates"][0]["content"]["parts"][0]["text"])

# ----------------- WORD BELGESİ OLUŞTURMA -----------------
def generate_word_report(report_date, dept, location, expert, sicil_no, findings):
    doc = Document()
    
    # Sayfa Boyutu: A4 Yatay
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    
    formatted_date = report_date.strftime("%d.%m.%Y")
    logo_exists = os.path.exists("logo.png")

    # ================= 1. KAPAK SAYFASI =================
    p_cov_logo = doc.add_paragraph()
    p_cov_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov_logo.paragraph_format.space_before = Pt(30)
    p_cov_logo.paragraph_format.space_after = Pt(20)
    
    if logo_exists:
        p_cov_logo.add_run().add_picture("logo.png", width=Inches(3.5)) # Logo büyütüldü
    else:
        r_l = p_cov_logo.add_run("[LOGO - logo.png]")
        r_l.bold = True

    # Fontlar tam 24 Punto (Orijinal ile birebir)
    add_cover_line(doc, "T.C.", 24, bold=True, space_after=6)
    add_cover_line(doc, "MERSİN BÜYÜKŞEHİR BELEDİYESİ", 24, bold=True, space_after=6)
    add_cover_line(doc, f"{dept.upper()}", 24, bold=True, space_after=12)
    add_cover_line(doc, f"Sicil No: {sicil_no}", 24, bold=False, space_after=40)
    add_cover_line(doc, "İSG Saha Gözetim Raporu", 24, bold=True, space_after=12)
    
    p_date = add_cover_line(doc, formatted_date, 24, bold=False, space_after=0)

    # Boş sayfa hatasını engellemek için page_break'i son yazının doğrudan içine gömüyoruz
    p_date.runs[0].add_break(docx.enum.text.WD_BREAK.PAGE)

    # ================= 2. GİRİŞ CÜMLESİ =================
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Tablo ile üst antete yapışmaması için boşluklar
    p_intro.paragraph_format.space_before = Pt(12) 
    p_intro.paragraph_format.space_after = Pt(12)

    r_in1 = p_intro.add_run(f"{formatted_date} tarihinde {dept}’na Bağlı ")
    r_in2 = p_intro.add_run(f"{location} ")
    r_in2.bold = True
    r_in3 = p_intro.add_run(
        "İş Güvenliği Uzmanı tarafından yapılan saha denetiminde, tespit edilen hususlara ilişkin "
        "“Tespit / Uygunsuzluk ve Öneriler” aşağıda belirtilmiştir."
    )
    for r in [r_in1, r_in2, r_in3]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    # ================= 3. ÜST BİLGİ (HEADER) AYARI =================
    section.different_first_page_header_footer = True
    header = section.header
    header_table = header.add_table(rows=2, cols=3, width=Inches(10.1))
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    set_table_borders(header_table, "000000")
    
    col_w = [Inches(1.8), Inches(6.1), Inches(2.2)]
    for row in header_table.rows:
        for i, w in enumerate(col_w):
            row.cells[i].width = w

    cell_h_logo = header_table.cell(0, 0)
    cell_h_logo.merge(header_table.cell(1, 0))
    cell_h_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_h_logo = cell_h_logo.paragraphs[0]
    p_h_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_exists:
        p_h_logo.add_run().add_picture("logo.png", width=Inches(1.3))
    else:
        p_h_logo.add_run("[LOGO]").bold = True

    cell_h_mid = header_table.cell(0, 1)
    cell_h_mid.merge(header_table.cell(1, 1))
    cell_h_mid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_h_mid = cell_h_mid.paragraphs[0]
    p_h_mid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    rh1 = p_h_mid.add_run("İŞ SAĞLIĞI VE GÜVENLİĞİ SAHA GÖZETİM RAPORU\n")
    rh2 = p_h_mid.add_run("MERSİN BÜYÜKŞEHİR BELEDİYESİ\n")
    rh3 = p_h_mid.add_run(f"{dept.upper()}\n")
    rh4 = p_h_mid.add_run(f"Sicil No : {sicil_no}")
    for r in [rh1, rh2, rh3, rh4]:
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    cell_h_date = header_table.cell(0, 2)
    p_hd = cell_h_date.paragraphs[0]
    r_h_dlbl = p_hd.add_run("Tarih: ")
    r_h_dlbl.bold = True
    r_h_dtxt = p_hd.add_run(formatted_date)
    for r in [r_h_dlbl, r_h_dtxt]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    cell_h_sec = header_table.cell(1, 2)
    p_hs = cell_h_sec.paragraphs[0]
    r_h_slbl = p_hs.add_run("Bölüm: ")
    r_h_slbl.bold = True
    r_h_stxt = p_hs.add_run(location)
    for r in [r_h_slbl, r_h_stxt]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    add_page_number_to_footer(section.footer)

    # ================= 4. BULGULAR TABLOSU =================
    findings_table = doc.add_table(rows=1, cols=2)
    findings_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    findings_table.autofit = False
    set_table_borders(findings_table, "000000")
    
    tbl_widths = [Inches(4.2), Inches(6.3)]
    
    hdr_cells = findings_table.rows[0].cells
    hdr_cells[0].width = tbl_widths[0]
    hdr_cells[1].width = tbl_widths[1]
    
    set_cell_background(hdr_cells[0], "F2F2F2")
    set_cell_background(hdr_cells[1], "F2F2F2")
    
    p_th0 = hdr_cells[0].paragraphs[0]
    p_th0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_th0 = p_th0.add_run("Fotoğraf")
    r_th0.bold = True
    r_th0.font.name = "Times New Roman"
    r_th0.font.size = Pt(12)
    
    p_th1 = hdr_cells[1].paragraphs[0]
    p_th1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_th1 = p_th1.add_run("Açıklamalar")
    r_th1.bold = True
    r_th1.font.name = "Times New Roman"
    r_th1.font.size = Pt(12)

    for f in findings:
        row_cells = findings_table.add_row().cells
        row_cells[0].width = tbl_widths[0]
        row_cells[1].width = tbl_widths[1]
        
        set_cell_margins(row_cells[0], top=120, bottom=120, left=100, right=100)
        set_cell_margins(row_cells[1], top=120, bottom=120, left=160, right=140)

        cell_img = row_cells[0]
        cell_img.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_img = cell_img.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if f["photos"]:
            num_p = len(f["photos"])
            calc_w = Inches(3.9 / num_p) if num_p > 1 else Inches(3.2)
            calc_w = min(calc_w, Inches(1.8)) if num_p > 1 else calc_w
            
            for p_idx, p_bytes in enumerate(f["photos"]):
                img_stream = io.BytesIO(p_bytes)
                p_img.add_run().add_picture(img_stream, width=calc_w)
                if p_idx < num_p - 1:
                    p_img.add_run("  ") 
        else:
            r_no = p_img.add_run("Görsel Eklenmedi")
            r_no.font.name = "Times New Roman"
            r_no.font.size = Pt(11)

        cell_desc = row_cells[1]
        cell_desc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        
        p_f = cell_desc.paragraphs[0]
        p_f.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_f_lbl = p_f.add_run("Tespit / Uygunsuzluk : ")
        r_f_lbl.bold = True
        r_f_lbl.underline = True
        r_f_txt = p_f.add_run(f["finding"])
        for r in [r_f_lbl, r_f_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            
        p_s = cell_desc.add_paragraph()
        p_s.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_s_lbl = p_s.add_run("Öneri: ")
        r_s_lbl.bold = True
        r_s_lbl.underline = True
        r_s_txt = p_s.add_run(f["suggestion"])
        for r in [r_s_lbl, r_s_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            
        p_l_head = cell_desc.add_paragraph()
        r_l_lbl = p_l_head.add_run("Yasal Dayanak:")
        r_l_lbl.bold = True
        r_l_lbl.underline = True
        r_l_lbl.font.name = "Times New Roman"
        r_l_lbl.font.size = Pt(12)
        
        legal_lines = [l.strip() for l in f["legal"].split("\n") if l.strip()]
        for line in legal_lines:
            p_bullet = cell_desc.add_paragraph(style='List Paragraph') 
            r_bullet = p_bullet.add_run(line.replace("•", "").strip())
            r_bullet.font.name = "Times New Roman"
            r_bullet.font.size = Pt(12)
            
        p_r = cell_desc.add_paragraph()
        r_r_lbl = p_r.add_run("Sorumlular: ")
        r_r_lbl.bold = True
        r_r_lbl.underline = True
        r_r_txt = p_r.add_run(f["responsible"])
        for r in [r_r_lbl, r_r_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            
        p_t = cell_desc.add_paragraph()
        r_t_lbl = p_t.add_run("Termin: ")
        r_t_lbl.bold = True
        r_t_lbl.underline = True
        r_t_txt = p_t.add_run(f["termin"])
        for r in [r_t_lbl, r_t_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)

    # ================= 5. KAPANIS CÜMLESİ VE İMZA BLOĞU =================
    # Son satırdan önceki boşluğu minimize ediyoruz
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(6)

    p_close = doc.add_paragraph()
    p_close.alignment = WD_ALIGN_PARAGRAPH.LEFT
    count = len(findings)
    r_close = p_close.add_run(
        f"İş bu rapor {count} madde tehlike ve buna bağlı oluşabilecek riskler ve alınması gereken önlemler olmak üzere "
        f"toplam {count} sayfadan ibarettir. ({formatted_date})"
    )
    r_close.font.name = "Times New Roman"
    r_close.font.size = Pt(12)
    
    # İmza Bloğu (Gereksiz paragraflar silindi, line spacing düzeltildi)
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.paragraph_format.space_before = Pt(24) # İmza aralığı
    
    r_s1 = p_sign.add_run("İş Sağlığı ve Güvenliği Uzmanı\n\n\n")
    r_s1.bold = True
    r_s2 = p_sign.add_run(f"{expert}\n")
    r_s2.bold = True
    r_s3 = p_sign.add_run("İmza")
    
    for r in [r_s1, r_s2, r_s3]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ----------------- STREAMLIT ÖN YÜZ -----------------
st.title("🛡️ İSG Saha Gözetim Raporu")

with st.container():
    st.subheader("📋 Rapor Üst Bilgileri")
    c1, c2 = st.columns(2)
    with c1:
        report_date = st.date_input("Rapor Tarihi", datetime.date.today())
        dept = st.selectbox(
            "Daire Başkanlığı",
            ["Veteriner İşleri Dairesi Başkanlığı", "Tarımsal Hizmetler Dairesi Başkanlığı"]
        )
    with c2:
        location = st.text_input("Lokasyon / İşletme", placeholder="Örn: Mut Mezbahası")
        expert = st.text_input("İSG Uzmanı", value="Kemal ŞAHİN")

    with st.expander("Ek Bilgi (Sicil Numarası)"):
        sicil_no = st.text_input("Sicil No", value="1 3900 01 01 1170275 033 12 53")

st.divider()

st.subheader("🔍 Tespit / Uygunsuzluk Ekle")

col_note, col_btn = st.columns([3.5, 1.5])
with col_note:
    raw_note = st.text_input("Saha Notunuz (Hızlı Giriş)", placeholder="Örn: Priz kapağı kırık, kablolar dağınık")
with col_btn:
    st.write("")
    st.write("")
    ai_btn = st.button("✨ Yapay Zeka ile Düzenle", use_container_width=True)

if "current_tespit" not in st.session_state:
    st.session_state.current_tespit = ""
if "current_oneri" not in st.session_state:
    st.session_state.current_oneri = ""
if "current_yasal" not in st.session_state:
    st.session_state.current_yasal = ""

if ai_btn:
    if not active_api_key:
        st.error("Lütfen sol taraftan Gemini API anahtarınızı giriniz.")
    elif not raw_note.strip():
        st.warning("Lütfen kısa bir saha gözlem notu girin.")
    else:
        with st.spinner("İSG standartlarına göre düzenleniyor..."):
            try:
                res = call_gemini_ai(raw_note, active_api_key)
                st.session_state.current_tespit = res.get("tespit", "")
                st.session_state.current_oneri = res.get("oneri", "")
                st.session_state.current_yasal = res.get("yasal", "")
                st.success("Yapay zeka düzenledi!")
            except Exception as e:
                st.error(f"Hata: {str(e)}")

finding_input = st.text_area("Tespit / Uygunsuzluk", value=st.session_state.current_tespit, height=85)
suggestion_input = st.text_area("Öneri / Önlem", value=st.session_state.current_oneri, height=85)

col_l, col_r, col_t = st.columns([2, 1.2, 1])
with col_l:
    legal_input = st.text_area("Yasal Dayanak (Satır Satır)", value=st.session_state.current_yasal, height=80)
with col_r:
    responsible_input = st.text_input("Sorumlular", value=dept)
with col_t:
    termin_input = st.text_input("Termin", value="İvedilikle")

uploaded_photos = st.file_uploader(
    "Fotoğraflar (Birden fazla seçebilirsiniz)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key=f"photo_uploader_{st.session_state.uploader_key}"
)

if uploaded_photos:
    st.caption("Yüklenen Fotoğraf Önizlemeleri:")
    st.markdown('<div class="preview-box">', unsafe_allow_html=True)
    cols = st.columns(min(len(uploaded_photos), 6))
    for idx, photo in enumerate(uploaded_photos):
        with cols[idx % 6]:
            st.image(photo, width=100)
    st.markdown('</div>', unsafe_allow_html=True)

if st.button("➕ Bu Maddeyi Listeye Kaydet", type="primary", use_container_width=True):
    if not finding_input.strip():
        st.warning("Tespit alanı boş bırakılamaz.")
    else:
        photo_bytes_list = [p.getvalue() for p in uploaded_photos] if uploaded_photos else []
        
        st.session_state.findings.append({
            "finding": finding_input.strip(),
            "suggestion": suggestion_input.strip(),
            "legal": legal_input.strip(),
            "responsible": responsible_input.strip(),
            "termin": termin_input.strip(),
            "photos": photo_bytes_list
        })
        
        st.session_state.current_tespit = ""
        st.session_state.current_oneri = ""
        st.session_state.current_yasal = ""
        
        st.session_state.uploader_key += 1
        st.success("Madde başarıyla listeye eklendi!")
        st.rerun()

st.divider()

st.subheader("📑 Eklenen Maddeler")

if not st.session_state.findings:
    st.info("Henüz eklenen madde yok.")
else:
    for idx, item in enumerate(st.session_state.findings):
        with st.expander(f"Madde {idx + 1}: {item['finding'][:70]}..."):
            st.markdown(f"**Tespit:** {item['finding']}")
            st.markdown(f"**Öneri:** {item['suggestion']}")
            st.markdown(f"**Yasal Dayanak:**\n{item['legal']}")
            st.caption(f"Fotoğraf Sayısı: {len(item['photos'])} | Sorumlu: {item['responsible']}")
            if st.button(f"🗑️ Sil", key=f"del_btn_{idx}"):
                st.session_state.findings.pop(idx)
                st.rerun()

    st.write("")
    loc_name = location.strip().replace(" ", "_") if location else "Saha"
    doc_bytes = generate_word_report(
        report_date=report_date,
        dept=dept,
        location=location or "..................",
        expert=expert,
        sicil_no=sicil_no,
        findings=st.session_state.findings
    )
    
    st.download_button(
        label="📥 YATAY WORD RAPORUNU İNDİR (.DOCX)",
        data=doc_bytes,
        file_name=f"Saha_Raporu-{loc_name}_{report_date.strftime('%d.%m.%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )
