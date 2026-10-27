# PDF İzin–Mesai Otomasyonu

Streamlit tabanlı web uygulaması. Birden fazla el yazılı PDF'yi görsel olarak OpenAI API üzerinden analiz eder, personel adı/sicil/tarih/saat ve İZİN-MESAİ ayrımını çıkarır; kullanıcıya düzenlenebilir tablo ve renkli Excel sunar.

## Çalıştırma

Python 3.10+ kurun. Bu klasörde terminal açıp:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tarayıcıda açılan sayfada OpenAI API anahtarını girin, PDF'leri yükleyin ve **PDF belgelerini oku** butonuna basın. API kullanımı ayrıca ücretlendirilebilir.

## İnternete yayınlama

Kodu özel bir GitHub deposuna yükleyip Streamlit Community Cloud veya uygun bir kurumsal sunucu üzerinden yayınlayabilirsiniz. Personel bilgileri içeren evraklar için kurum onayı, erişim kontrolü ve veri işleme koşulları gereklidir. Bu ZIP kendi başına canlı bir internet adresi değildir.

## Notlar

El yazısı ve kutucuk tanıma kesin değildir. Güven alanı düşükse doğrulayın. Uygulama sonuçları varsayılan olarak sunucuya kalıcı dosya olarak yazmaz. Harici API'ye belge görselleri gönderilir. Bu sürüm PDF'den talep çıkarmak içindir; PDKS karşılaştırması içermez.
