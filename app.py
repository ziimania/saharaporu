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

# ----------------- SAYFA KONFİGÜRASYONU -----------------
st.set_page_config(
    page_title="İSG Saha Gözetim Raporu Oluşturucu",
    page_icon="🛡️",
    layout="wide"
)

# Çerez Yöneticisi (Tarayıcıda saklama için)
cookie_controller = CookieController()

# Session State Başlatma
if "findings" not in st.session_state:
    st.session_state.findings = []
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

# Çerezden API anahtarını oku
saved_cookie_key = cookie_controller.get("gemini_api_key")
if saved_cookie_key and not st.session_state.api_key:
    st.session_state.api_key = saved_cookie_key

# ----------------- SOL KENAR ÇUBUĞU (API & AYARLAR) -----------------
with st.sidebar:
    st.header("⚙️ Sistem Ayarları")
    st.caption("API Anahtarınız buluta gitmez, yalnızca kendi tarayıcınızın çerezinde saklanır.")
    
    input_key = st.text_input(
        "Gemini API Anahtarı",
        type="password",
        value=st.session_state.api_key,
        help="Google AI Studio üzerinden aldığınız anahtarı giriniz. Tarayıcınız şifre kaydetmeyi teklif edebilir."
    )
    
    col_save, col_clear = st.columns(2)
    with col_save:
        if st.button("💾 Çereze Kaydet", use_container_width=True):
            if input_key:
                st.session_state.api_key = input_key
                cookie_controller.set("gemini_api_key", input_key, max_age=30*24*60*60) # 30 gün geçerli
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
    st.markdown("**Gözetim Özeti**")
    st.info(f"Eklenen Madde Sayısı: **{len(st.session_state.findings)}**")

# ----------------- YARDIMCI STİL FONKSİYONLARI -----------------
def set_cell_background(cell, hex_color):
    """Hücreye arka plan rengi verir."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    """Hücre içi kenar boşlukları (padding)."""
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

def set_table_borders(table, color="D3D3D3"):
    """Tabloya profesyonel ince kenarlık ekler."""
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
    
    data = response.json()
    import json
    return json.loads(data["candidates"][0]["content"]["parts"][0]["text"])

# ----------------- WORD BELGESİ ÜRETİMİ (BİREBİR ORİJİNAL FORMAT) -----------------
def generate_word_report(report_date, dept, location, expert, sicil_no, findings):
    doc = Document()
    
    # Sayfa Yapısı: A4 Yatay (Landscape)
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    
    # 1. BAŞLIK VE ANTET TABLOSU
    header_table = doc.add_table(rows=2, cols=3)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    set_table_borders(header_table, "000000")
    
    col_widths = [Inches(1.8), Inches(6.4), Inches(2.3)]
    for row in header_table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w
            
    # Sol Hücre (Logo alanı)
    cell_logo = header_table.cell(0, 0)
    cell_logo.merge(header_table.cell(1, 0))
    cell_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_logo = cell_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Logoyu web üzerinden çek veya metin koy
    try:
        logo_url = "https://mersinbuyuksehiras.com.tr/wp-content/uploads/2024/04/image-13.svg"
        # Word doğrudan SVG kabul etmediği için şık bir kurumsal başlık ikonu ekliyoruz
        r_logo = p_logo.add_run("MERSİN\nBÜYÜKŞEHİR\nBELEDİYESİ")
        r_logo.bold = True
        r_logo.font.name = "Times New Roman"
        r_logo.font.size = Pt(11)
        r_logo.font.color.rgb = RGBColor(0, 51, 153)
    except:
        p_logo.add_run("[LOGO]")

    # Orta Hücre (Kurum ve Daire Bilgisi)
    cell_center = header_table.cell(0, 1)
    cell_center.merge(header_table.cell(1, 1))
    cell_center.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_center = cell_center.paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    r1 = p_center.add_run("İŞ SAĞLIĞI VE GÜVENLİĞİ SAHA GÖZETİM RAPORU\n")
    r1.bold = True
    r2 = p_center.add_run("MERSİN BÜYÜKŞEHİR BELEDİYESİ\n")
    r2.bold = True
    r3 = p_center.add_run(f"{dept.upper()}\n")
    r3.bold = True
    r4 = p_center.add_run(f"Sicil No : {sicil_no}")
    
    for r in [r1, r2, r3, r4]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    # Sağ Üst Hücre (Tarih)
    cell_date = header_table.cell(0, 2)
    p_date = cell_date.paragraphs[0]
    r_d_label = p_date.add_run("Tarih: ")
    r_d_label.bold = True
    r_d_val = p_date.add_run(report_date.strftime("%d.%m.%Y"))
    for r in [r_d_label, r_d_val]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    # Sağ Alt Hücre (Bölüm / Lokasyon)
    cell_sec = header_table.cell(1, 2)
    p_sec = cell_sec.paragraphs[0]
    r_s_label = p_sec.add_run("Bölüm: ")
    r_s_label.bold = True
    r_s_val = p_sec.add_run(location)
    for r in [r_s_label, r_s_val]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    doc.add_paragraph() # Boşluk

    # 2. GİRİŞ PARAGRAFI
    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.line_spacing = 1.15
    p_intro.paragraph_format.space_after = Pt(10)
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    formatted_date = report_date.strftime("%d.%m.%Y")
    intro_run1 = p_intro.add_run(f"{formatted_date} tarihinde {dept}’na Bağlı ")
    intro_run2 = p_intro.add_run(f"{location} ")
    intro_run2.bold = True
    intro_run3 = p_intro.add_run(
        "İş Güvenliği Uzmanı tarafından yapılan saha denetiminde, tespit edilen hususlara ilişkin "
        "“Tespit / Uygunsuzluk ve Öneriler” aşağıda belirtilmiştir."
    )
    for r in [intro_run1, intro_run2, intro_run3]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    # 3. BULGULAR TABLOSU
    findings_table = doc.add_table(rows=1, cols=2)
    findings_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    findings_table.autofit = False
    set_table_borders(findings_table, "000000")
    
    f_widths = [Inches(4.2), Inches(6.3)]
    
    # Başlık Satırı
    hdr_cells = findings_table.rows[0].cells
    hdr_cells[0].width = f_widths[0]
    hdr_cells[1].width = f_widths[1]
    
    set_cell_background(hdr_cells[0], "F2F2F2")
    set_cell_background(hdr_cells[1], "F2F2F2")
    
    p_h0 = hdr_cells[0].paragraphs[0]
    p_h0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h0 = p_h0.add_run("Fotoğraf")
    r_h0.bold = True
    r_h0.font.name = "Times New Roman"
    r_h0.font.size = Pt(12)
    
    p_h1 = hdr_cells[1].paragraphs[0]
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h1 = p_h1.add_run("Açıklamalar")
    r_h1.bold = True
    r_h1.font.name = "Times New Roman"
    r_h1.font.size = Pt(12)
    
    # Madde Satırları
    for f in findings:
        row_cells = findings_table.add_row().cells
        row_cells[0].width = f_widths[0]
        row_cells[1].width = f_widths[1]
        
        set_cell_margins(row_cells[0], top=120, bottom=120, left=120, right=120)
        set_cell_margins(row_cells[1], top=120, bottom=120, left=180, right=140)
        
        # --- Sol Hücre: Fotoğraflar ---
        cell_img = row_cells[0]
        cell_img.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_img = cell_img.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if f["photos"]:
            for photo_bytes in f["photos"]:
                img_stream = io.BytesIO(photo_bytes)
                im = Image.open(img_stream)
                width, height = im.size
                
                # En/boy oranına göre optimize boyutlandırma
                img_stream.seek(0)
                if width > height:
                    p_img.add_run().add_picture(img_stream, width=Inches(3.8))
                else:
                    p_img.add_run().add_picture(img_stream, height=Inches(3.4))
                p_img.add_run("\n")
        else:
            r_no_img = p_img.add_run("Görsel Eklenmedi")
            r_no_img.font.name = "Times New Roman"
            r_no_img.font.size = Pt(11)

        # --- Sağ Hücre: Açıklamalar ---
        cell_desc = row_cells[1]
        cell_desc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        
        # Tespit
        p_f = cell_desc.paragraphs[0]
        p_f.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_f_lbl = p_f.add_run("Tespit / Uygunsuzluk : ")
        r_f_lbl.bold = True
        r_f_txt = p_f.add_run(f["finding"])
        for r in [r_f_lbl, r_f_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            
        # Öneri
        p_s = cell_desc.add_paragraph()
        p_s.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_s_lbl = p_s.add_run("Öneri: ")
        r_s_lbl.bold = True
        r_s_txt = p_s.add_run(f["suggestion"])
        for r in [r_s_lbl, r_s_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            
        # Yasal Dayanak
        p_l_head = cell_desc.add_paragraph()
        r_l_lbl = p_l_head.add_run("Yasal Dayanak:")
        r_l_lbl.bold = True
        r_l_lbl.font.name = "Times New Roman"
        r_l_lbl.font.size = Pt(12)
        
        # Madde imli kanunlar
        legal_lines = [l.strip() for l in f["legal"].split("\n") if l.strip()]
        for line in legal_lines:
            p_bullet = cell_desc.add_paragraph(style='List Paragraph')
            p_bullet.paragraph_format.left_indent = Inches(0.25)
            p_bullet.paragraph_format.space_after = Pt(2)
            r_bullet = p_bullet.add_run(line)
            r_bullet.font.name = "Times New Roman"
            r_bullet.font.size = Pt(12)
            
        # Sorumlular
        p_r = cell_desc.add_paragraph()
        r_r_lbl = p_r.add_run("Sorumlular: ")
        r_r_lbl.bold = True
        r_r_txt = p_r.add_run(f["responsible"])
        for r in [r_r_lbl, r_r_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            
        # Termin
        p_t = cell_desc.add_paragraph()
        r_t_lbl = p_t.add_run("Termin: ")
        r_t_lbl.bold = True
        r_t_txt = p_t.add_run(f["termin"])
        for r in [r_t_lbl, r_t_txt]:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)

    # 4. KAPANIS CÜMLESİ VE İMZA BLOĞU
    doc.add_paragraph()
    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.LEFT
    count = len(findings)
    r_foot = p_footer.add_run(
        f"İş bu rapor {count} madde tehlike ve buna bağlı oluşabilecek riskler ve alınması gereken önlemler olmak üzere "
        f"toplam {count} sayfadan ibarettir. ({formatted_date})"
    )
    r_foot.font.name = "Times New Roman"
    r_foot.font.size = Pt(12)
    
    # İmza Bloğu (Sağ Alta Hizalı)
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_s1 = p_sign.add_run("İş Sağlığı ve Güvenliği Uzmanı\n\n\n")
    r_s1.bold = True
    r_s2 = p_sign.add_run(f"{expert}\n")
    r_s2.bold = True
    r_s3 = p_sign.add_run("İmza")
    for r in [r_s1, r_s2, r_s3]:
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    # Belgeyi bayt akışı olarak döndür
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ----------------- ANA STREAMLIT ARAYÜZÜ -----------------
st.title("🛡️ İSG Saha Gözetim Raporu Girişi")

# 1. Rapor Üst Bilgileri
with st.container():
    st.subheader("📋 Rapor Üst Bilgileri")
    c1, c2, c3, c4 = st.columns([1.5, 2.5, 2.5, 1.5])
    
    with c1:
        report_date = st.date_input("Rapor Tarihi", datetime.date.today())
    with c2:
        dept_options = [
            "Veteriner İşleri Dairesi Başkanlığı",
            "Tarımsal Hizmetler Dairesi Başkanlığı"
        ]
        dept = st.selectbox("Daire Başkanlığı", dept_options)
    with c3:
        location = st.text_input("Lokasyon / İşletme", placeholder="Örn: Mut Mezbahası")
    with c4:
        expert = st.text_input("İSG Uzmanı", value="Kemal ŞAHİN")

    # Sicil No (Opsiyonel / Özelleştirilebilir)
    default_sicil = "1 3900 01 01 1170275 033 11 53"
    with st.expander("Ek Üst Bilgi (Sicil Numarası)"):
        sicil_no = st.text_input("Sicil No", value=default_sicil)

st.divider()

# 2. Tespit ve Uygunsuzluk Giriş Alanı
st.subheader("🔍 Tespit / Uygunsuzluk Ekle")

col_note, col_btn = st.columns([4, 1])
with col_note:
    raw_note = st.text_input(
        "Saha Notunuz (Hızlı Giriş)", 
        placeholder="Örn: Priz kapağı kırık, kablolar dışarıda dağınık duruyor"
    )
with col_btn:
    st.write("") # Hizalama boşluğu
    st.write("")
    ai_btn = st.button("✨ Yapay Zeka ile Düzenle", use_container_width=True)

# Oturumda form alanlarını tutmak için
if "current_tespit" not in st.session_state:
    st.session_state.current_tespit = ""
if "current_oneri" not in st.session_state:
    st.session_state.current_oneri = ""
if "current_yasal" not in st.session_state:
    st.session_state.current_yasal = ""

if ai_btn:
    if not active_api_key:
        st.error("Lütfen önce sol taraftan Gemini API anahtarınızı giriniz.")
    elif not raw_note:
        st.warning("Lütfen düzenlenecek kısa bir saha notu yazınız.")
    else:
        with st.spinner("Yapay zeka İSG standartlarına göre raporlaştırıyor..."):
            try:
                ai_res = call_gemini_ai(raw_note, active_api_key)
                st.session_state.current_tespit = ai_res.get("tespit", "")
                st.session_state.current_oneri = ai_res.get("oneri", "")
                st.session_state.current_yasal = ai_res.get("yasal", "")
                st.success("Yapay zeka başarıyla düzenledi! Aşağıdaki alanları kontrol edebilirsiniz.")
            except Exception as e:
                st.error(f"Hata oluştu: {str(e)}")

# Form Alanları
col_f1, col_f2 = st.columns(2)
with col_f1:
    finding_input = st.text_area(
        "Tespit / Uygunsuzluk", 
        value=st.session_state.current_tespit, 
        height=100,
        help="Cümlenin '... tespit edilmiştir.' ile bittiğinden emin olunuz."
    )
with col_f2:
    suggestion_input = st.text_area(
        "Öneri / Önlem", 
        value=st.session_state.current_oneri, 
        height=100
    )

col_f3, col_f4, col_f5 = st.columns([2, 1, 1])
with col_f3:
    legal_input = st.text_area(
        "Yasal Dayanak (Her satıra bir yönetmelik/kanun)", 
        value=st.session_state.current_yasal, 
        height=85
    )
with col_f4:
    responsible_input = st.text_input("Sorumlu", value=dept)
with col_f5:
    termin_input = st.text_input("Termin", value="İvedilikle")

# Fotoğraf Yükleyici
uploaded_photos = st.file_uploader(
    "Madde Fotoğrafları (Birden fazla seçebilirsiniz)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_photos:
    st.caption("Yüklenen Fotoğraf Önizlemeleri:")
    preview_cols = st.columns(min(len(uploaded_photos), 5))
    for idx, photo in enumerate(uploaded_photos):
        with preview_cols[idx % 5]:
            st.image(photo, use_container_width=True)

# Listeye Kaydet Butonu
if st.button("➕ Bu Maddeyi Rapor Listesine Kaydet", type="primary", use_container_width=True):
    if not finding_input.strip():
        st.warning("Tespit alanı boş bırakılamaz.")
    else:
        # Fotoğrafları bayt olarak hafızaya al
        photo_bytes_list = []
        if uploaded_photos:
            for p in uploaded_photos:
                photo_bytes_list.append(p.getvalue())
                
        st.session_state.findings.append({
            "finding": finding_input.strip(),
            "suggestion": suggestion_input.strip(),
            "legal": legal_input.strip(),
            "responsible": responsible_input.strip(),
            "termin": termin_input.strip(),
            "photos": photo_bytes_list
        })
        
        # Alanları sıfırla
        st.session_state.current_tespit = ""
        st.session_state.current_oneri = ""
        st.session_state.current_yasal = ""
        st.success("Madde rapora eklendi!")
        st.rerun()

st.divider()

# 3. Eklenen Maddeler Listesi
st.subheader("📑 Rapora Eklenen Maddeler")

if not st.session_state.findings:
    st.info("Henüz eklenmiş bir madde bulunmuyor. Yukarıdaki formu kullanarak madde ekleyebilirsiniz.")
else:
    for idx, item in enumerate(st.session_state.findings):
        with st.expander(f"Madde {idx + 1}: {item['finding'][:80]}..."):
            st.markdown(f"**Tespit:** {item['finding']}")
            st.markdown(f"**Öneri:** {item['suggestion']}")
            st.markdown(f"**Yasal Dayanak:**\n{item['legal']}")
            st.markdown(f"**Sorumlu:** {item['responsible']} | **Termin:** {item['termin']}")
            st.caption(f"Ekli Fotoğraf Sayısı: {len(item['photos'])}")
            
            if st.button(f"🗑️ Madde {idx + 1}'i Sil", key=f"del_{idx}"):
                st.session_state.findings.pop(idx)
                st.rerun()

    # 4. Word İndirme Butonu
    st.write("")
    loc_clean = location.strip().replace(" ", "_") if location else "Saha"
    date_str = report_date.strftime("%d.%m.%Y")
    file_name = f"Saha_Raporu-{loc_clean}_{date_str}.docx"
    
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
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )