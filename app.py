import io
import re
from datetime import datetime
import streamlit as st
import pandas as pd
import fitz
from PIL import Image
import pytesseract
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='PDF İzin–Mesai Otomasyonu', page_icon='📄', layout='wide')
st.title('📄 PDF İzin–Mesai Otomasyonu — Ücretsiz Sürüm')
st.info('API anahtarı gerekmez. PDF metni ve ücretsiz OCR kullanılır. El yazısı, ad ve tarihler mutlaka kontrol edilmelidir.')
st.warning('Bu sürüm yalnız PDF taleplerini listeler; PDKS 131 kayıtlarıyla otomatik karşılaştırma yapmaz. Gerçek personel evraklarını yüklemeden önce kurumunuzun veri güvenliği kurallarını kontrol edin.')

COLS = ['Dosya','Ad Soyad','Sicil No','Belge Türü','Talep','Tarih','Başlangıç','Bitiş','Açıklama','Okuma Güveni']
DATE_RE = re.compile(r'(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](20\d{2})(?!\d)')
TIME_RE = re.compile(r'(?<!\d)([01]?\d|2[0-3])[:.]([0-5]\d)(?!\d)')

def dates_in(text):
    dates = []
    for d,m,y in DATE_RE.findall(text):
        try:
            dt = datetime(int(y), int(m), int(d))
            val = dt.strftime('%d.%m.%Y')
            if val not in dates:
                dates.append(val)
        except ValueError:
            pass
    return dates

def read_pdf(upload):
    doc = fitz.open(stream=upload.getvalue(), filetype='pdf')
    if len(doc) > 20:
        raise ValueError('PDF 20 sayfadan uzun; lütfen bölün.')
    parts=[]
    for page in doc:
        direct = page.get_text('text').strip()
        if len(direct) >= 40:
            parts.append((direct,'PDF metni'))
        else:
            pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
            img=Image.open(io.BytesIO(pix.tobytes('png')))
            try:
                ocr=pytesseract.image_to_string(img, lang='tur+eng', config='--psm 3')
            except pytesseract.TesseractError:
                ocr=pytesseract.image_to_string(img, lang='eng', config='--psm 3')
            parts.append((ocr,'OCR (kontrol gerekli)'))
    return parts

def analyze(filename, page_number, text, method):
    found=dates_in(text)
    times=[f'{h.zfill(2)}:{m}' for h,m in TIME_RE.findall(text)]
    name=''
    match=re.search(r'(?:Ad[ıi]\s*Soyad[ıi]|Ad[ıi]\s*ve\s*Soyad[ıi])\s*[:：-]?\s*([^\n]{3,55})',text,re.I)
    if match: name=match.group(1).strip()
    sicil=''
    match=re.search(r'Sicil\s*(?:No|Numaras[ıi])?\s*[:：-]?\s*(\d{3,12})',text,re.I)
    if match: sicil=match.group(1)
    lower=text.casefold()
    if 'denkleştirme izni' in lower or 'denklestirme izni' in lower:
        talep='KONTROL EDİLEMEDİ'  # checkbox selection cannot be reliably inferred
    else:
        talep='KONTROL EDİLEMEDİ'
    note=f'Sayfa {page_number}; {method}. Talep türünü, adı ve tarihleri evraktan doğrulayın. Metin örneği: {text[:240].replace(chr(10)," ")}'
    base={'Dosya':filename,'Ad Soyad':name,'Sicil No':sicil,'Belge Türü':'','Talep':talep,'Başlangıç':times[0] if times else '', 'Bitiş':times[1] if len(times)>1 else '', 'Açıklama':note,'Okuma Güveni':'DÜŞÜK' if method.startswith('OCR') else 'ORTA'}
    return [{**base,'Tarih':date} for date in (found or [''])]

files=st.file_uploader('PDF evraklarını seçin',type=['pdf'],accept_multiple_files=True)
if 'results' not in st.session_state:
    st.session_state.results=[]
if st.button('🔍 PDF belgelerini oku',type='primary',disabled=not files):
    result=[]
    progress=st.progress(0)
    for idx, f in enumerate(files):
        try:
            pages=read_pdf(f)
            for i,(text,method) in enumerate(pages,1):
                result.extend(analyze(f.name,i,text,method))
        except Exception as exc:
            result.append({'Dosya':f.name,'Talep':'KONTROL EDİLEMEDİ','Açıklama':f'Okuma hatası: {exc}','Okuma Güveni':'DÜŞÜK'})
        progress.progress((idx+1)/len(files))
    st.session_state.results=result

if st.session_state.results:
    st.subheader('Sonuçlar — evrakla karşılaştırıp düzeltin')
    df=pd.DataFrame(st.session_state.results).reindex(columns=COLS).fillna('')
    edited=st.data_editor(df,hide_index=True,num_rows='dynamic',use_container_width=True,
        column_config={'Talep':st.column_config.SelectboxColumn('Talep',options=['İZİN','MESAİ','KONTROL EDİLEMEDİ'],required=True)})
    def export_xlsx(frame):
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as writer:
            frame.to_excel(writer,index=False,sheet_name='PDF Talepleri')
            ws=writer.sheets['PDF Talepleri']
            ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.fill=PatternFill('solid',fgColor='17365D')
                cell.font=Font(bold=True,color='FFFFFF')
            for row in ws.iter_rows(min_row=2):
                color={'İZİN':'DFF0D8','MESAİ':'DDEBF7','KONTROL EDİLEMEDİ':'FFF2CC'}.get(row[4].value,'FFFFFF')
                for cell in row:
                    cell.fill=PatternFill('solid',fgColor=color)
                    cell.alignment=Alignment(vertical='top',wrap_text=True)
            for i,width in enumerate([35,25,15,25,24,16,15,15,75,18],1):
                ws.column_dimensions[get_column_letter(i)].width=width
        return out.getvalue()
    st.download_button('⬇️ Renkli Excel raporunu indir',data=export_xlsx(edited),file_name='PDF_Izin_Mesai_Raporu.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
    st.caption('Yeşil: İZİN • Mavi: MESAİ • Sarı: KONTROL EDİLEMEDİ. Otomatik okunan bilgiler kesin doğrulanmış sayılmaz.')
