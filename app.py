import io, os, json, base64
import streamlit as st
import pandas as pd
import fitz
from openai import OpenAI
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='PDF İzin–Mesai Otomasyonu',page_icon='📄',layout='wide')
st.title('📄 PDF İzin–Mesai Otomasyonu')
st.caption('El yazılı evrakları okuyun, izin ve mesai taleplerini renkli Excel raporuna dönüştürün.')
st.warning('Personel verileri hassastır. Kurumunuzun veri güvenliği kurallarına uygun şekilde kullanın. Belgeler analiz için seçtiğiniz yapay zekâ sağlayıcısına gönderilir.')
api_key=st.text_input('OpenAI API anahtarı',type='password',value=os.getenv('OPENAI_API_KEY',''),help='API kullanımı ücretli olabilir. Anahtar bu uygulamada dosyaya kaydedilmez.')
files=st.file_uploader('PDF evraklarını seçin (birden fazla olabilir)',type=['pdf'],accept_multiple_files=True)

PROMPT='''Türkçe personel evrakını dikkatlice incele. El yazısını, işaretli kutuları ve açıklamaları birlikte değerlendir. "Fazla çalışma" başlığı tek başına mesai ücreti istendiği anlamına gelmez. Denkleştirme izni işaretliyse İZİN; mesai ücreti işaretliyse MESAİ; ikisi de kesin değilse KONTROL EDİLEMEDİ. Denkleştirme izni dilekçesi İZİN'dir. Birden fazla ayrı talep varsa ayrı kayıt üret. Bilgi uydurma. SADECE geçerli JSON nesnesi döndür: {"kayitlar":[{"ad_soyad":"","sicil_no":"","belge_turu":"","talep":"İZİN|MESAİ|KONTROL EDİLEMEDİ","tarih":"GG.AA.YYYY veya boş","baslangic":"SS:DD veya boş","bitis":"SS:DD veya boş","aciklama":"","guven":"YÜKSEK|ORTA|DÜŞÜK"}]}. İlgili birden fazla tarih varsa ayrı kayıtlar oluştur. Bir belge okunamıyorsa kontrol edilemedi kaydı döndür.'''

def parse_pdf(data):
    doc=fitz.open(stream=data,filetype='pdf')
    imgs=[]
    for page in doc:
        pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
        imgs.append('data:image/png;base64,'+base64.b64encode(pix.tobytes('png')).decode())
    return imgs

def extract(client,images):
    content=[{'type':'text','text':PROMPT}]+[{'type':'image_url','image_url':{'url':url}} for url in images]
    r=client.chat.completions.create(model='gpt-4o',messages=[{'role':'user','content':content}],response_format={'type':'json_object'},temperature=0)
    return json.loads(r.choices[0].message.content)['kayitlar']

COLUMNS=['Dosya','Ad Soyad','Sicil No','Belge Türü','Talep','Tarih','Başlangıç','Bitiş','Açıklama','Okuma Güveni']
if 'results' not in st.session_state: st.session_state.results=[]
if st.button('🔍 PDF belgelerini oku',type='primary',disabled=not(files and api_key)):
    client=OpenAI(api_key=api_key)
    results=[]
    progress=st.progress(0)
    for i,f in enumerate(files):
        try:
            pages=parse_pdf(f.getvalue())
            if len(pages)>20: raise ValueError('20 sayfadan uzun dosya: lütfen bölün.')
            for item in extract(client,pages):
                results.append({'Dosya':f.name,'Ad Soyad':item.get('ad_soyad',''),'Sicil No':item.get('sicil_no',''),'Belge Türü':item.get('belge_turu',''),'Talep':item.get('talep','KONTROL EDİLEMEDİ'),'Tarih':item.get('tarih',''),'Başlangıç':item.get('baslangic',''),'Bitiş':item.get('bitis',''),'Açıklama':item.get('aciklama',''),'Okuma Güveni':item.get('guven','DÜŞÜK')})
        except Exception as e:
            results.append({'Dosya':f.name,'Talep':'KONTROL EDİLEMEDİ','Açıklama':f'Okuma hatası: {str(e)}','Okuma Güveni':'DÜŞÜK'})
        progress.progress((i+1)/len(files))
    st.session_state.results=results

if st.session_state.results:
    st.subheader('Sonuçlar — indirmeden önce kontrol edin')
    df=pd.DataFrame(st.session_state.results).reindex(columns=COLUMNS).fillna('')
    edited=st.data_editor(df,num_rows='dynamic',use_container_width=True,hide_index=True)
    def to_xlsx(frame):
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as writer:
            frame.to_excel(writer,index=False,sheet_name='PDF Talepleri')
            ws=writer.sheets['PDF Talepleri']; ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.fill=PatternFill('solid',fgColor='17365D');cell.font=Font(bold=True,color='FFFFFF')
            for row in ws.iter_rows(min_row=2):
                talep=row[4].value
                color={'İZİN':'DFF0D8','MESAİ':'DDEBF7','KONTROL EDİLEMEDİ':'FFF2CC'}.get(talep,'FFFFFF')
                for c in row: c.fill=PatternFill('solid',fgColor=color);c.alignment=Alignment(vertical='top',wrap_text=True)
            widths=[35,25,15,30,23,16,15,15,60,18]
            for i,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(i)].width=w
        return out.getvalue()
    st.download_button('⬇️ Renkli Excel raporunu indir',data=to_xlsx(edited),file_name='PDF_Izin_Mesai_Raporu.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
    st.caption('Yeşil: İZİN • Mavi: MESAİ • Sarı: KONTROL EDİLEMEDİ. El yazısı sonuçlarını kontrol edip tabloda düzeltebilirsiniz.')
