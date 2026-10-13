import os
import feedparser
import urllib.parse
import pandas as pd
import re
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import google.generativeai as genai
import psycopg2 
import sys

# --- 1. AYARLAR VE ÇEVRE DEĞİŞKENLERİ ---
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL") 
GONDEREN_MAIL = os.environ.get("GONDEREN_MAIL")
GONDEREN_SIFRE = os.environ.get("GONDEREN_SIFRE")
ALICI_MAILLER = ["muratfirar@hotmail.com.tr", "muratfirar@isleasing.com.tr"] # Raporun gideceği mail adresini buraya yaz

# Eksik Secret kontrolü
if not all([GEMINI_API_KEY, SUPABASE_URL, GONDEREN_MAIL, GONDEREN_SIFRE]):
    print("KRİTİK HATA: GitHub Secrets (Şifreler) eksik veya okunmuyor!")
    sys.exit(1)

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-3.5-flash-latest')

ANAHTAR_KELIMELER = ["konkordato", "iflas", "dava", "haciz", "grev", "ceza", "zarar", "temerrüt"]
NEGATIF_KELIMELER = {"Genel": ["bedava", "promosyon", "reklam", "sponsor", "indirim", "dizi"]}

def tr_lower(text):
    if not isinstance(text, str): return ""
    return text.replace("&quot;", '"').replace("&#39;", "'").replace("&nbsp;", " ").replace("İ", "i").replace("I", "ı").lower()

# --- 2. SUPABASE BAĞLANTISI ---
try:
    conn = psycopg2.connect(SUPABASE_URL)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS risk_haberleri (
            id SERIAL PRIMARY KEY,
            firma TEXT, kaynak TEXT, tarih TEXT, baslik TEXT, link TEXT UNIQUE, llm_ozet TEXT
        )
    ''')
    conn.commit()
    print("Veritabanı bağlantısı başarılı.")
except Exception as e:
    print(f"KRİTİK HATA: Supabase bağlantısı kurulamadı. URL'yi kontrol edin. Detay: {e}")
    sys.exit(1)

# --- 3. EXCEL OKUMA ---
try:
    df_firmalar = pd.read_excel('firmalar.xlsx')
    print(f"Excel'den {len(df_firmalar)} satır başarıyla okundu.")
except Exception as e:
    print(f"KRİTİK HATA: firmalar.xlsx dosyası okunamadı. Depoda yüklü olduğundan emin olun. Detay: {e}")
    sys.exit(1)

raporlanacak_haberler = []

# --- 4. HABER TARAMA ---
for index, row in df_firmalar.iterrows():
    # Excel sütun adlarınıza dikkat edin. Burada 'firma_unvan' ve 'haber_unvan' aranmaktadır.
    try:
        firma_tam_ad = str(row['firma_unvan']).strip()
        haber_unvan = str(row['haber_unvan']).strip()
    except KeyError as e:
        print(f"KRİTİK HATA: Excel'de gerekli sütun bulunamadı. Lütfen sütun adlarının 'firma_unvan' ve 'haber_unvan' olduğundan emin olun. Detay: {e}")
        sys.exit(1)
    
    keyword_query = " OR ".join(ANAHTAR_KELIMELER)
    sorgu = f'"{haber_unvan}" AND ({keyword_query})'
    sorgu_encoded = urllib.parse.quote(sorgu)
    rss_url = f"https://news.google.com/rss/search?q={sorgu_encoded}&hl=tr&gl=TR"
    
    try:
        feed = feedparser.parse(rss_url)
        firma_negatifleri = NEGATIF_KELIMELER.get("Genel", [])
        
        for entry in feed.entries:
            baslik = entry.title
            link = entry.link
            yayim_tarihi = entry.published[5:16] if hasattr(entry, 'published') else "Bilinmeyen Tarih"
            kaynak_adi = entry.source.title if hasattr(entry, 'source') else "Google News"
            baslik_lower = tr_lower(baslik)
            
            if any(negatif in baslik_lower for negatif in firma_negatifleri):
                continue
                
            eslesenler = [kw for kw in ANAHTAR_KELIMELER if re.search(rf"(?:^|\W){kw}", baslik_lower)]
            
            if eslesenler:
                cursor.execute("SELECT id FROM risk_haberleri WHERE link=%s", (link,))
                if cursor.fetchone() is not None:
                    continue 
                
                prompt = f"""
                Sen bir kurumsal kredi risk analistisin. 
                Aşağıdaki haber metni '{firma_tam_ad}' ile doğrudan ilgili mi ve finansal/hukuki bir risk içeriyor mu? 
                İlgiliyse 2 cümlelik bir yönetici özeti yaz. Önemsiz bir bahsedilmeyse veya ilgisizse sadece 'RED' yaz.
                Haber: {baslik}
                """
                try:
                    llm_yaniti = model.generate_content(prompt).text.strip()
                except Exception as e:
                    print(f"Gemini API Hatası: {e}")
                    llm_yaniti = "RED"
                
                cursor.execute('''
                    INSERT INTO risk_haberleri (firma, kaynak, tarih, baslik, link, llm_ozet)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (link) DO NOTHING
                ''', (firma_tam_ad, kaynak_adi, yayim_tarihi, baslik, link, llm_yaniti))
                conn.commit()
                
                if "RED" not in llm_yaniti.upper():
                    raporlanacak_haberler.append({
                        "Firma": firma_tam_ad, "Başlık": baslik, "Özet": llm_yaniti, "Bağlantı": link
                    })
                    
    except Exception as e:
        print(f"[{haber_unvan}] RSS Tarama Hatası: {e}")
        
    time.sleep(2)

cursor.close()
conn.close()

# --- 5. MAİL GÖNDERİMİ ---
if raporlanacak_haberler:
    df_sonuclar = pd.DataFrame(raporlanacak_haberler)
    html_tablo = df_sonuclar.to_html(escape=False, index=False, formatters={'Bağlantı': lambda val: f'<a href="{val}">Haberi Oku</a>'})
    html_mail_icerik = f"<h2>Günlük Risk Raporu</h2>{html_tablo}"
    
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"⚠️ Erken Uyarı: {len(df_sonuclar)} Yeni Risk Haberi"
        msg["From"] = GONDEREN_MAIL
        msg["To"] = ", ".join(ALICI_MAILLER)
        msg.attach(MIMEText(html_mail_icerik, "html"))
        
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(GONDEREN_MAIL, GONDEREN_SIFRE)
        server.sendmail(GONDEREN_MAIL, ALICI_MAILLER, msg.as_string())
        server.quit()
        print("Mail başarıyla gönderildi!")
    except Exception as e:
        print(f"KRİTİK HATA: Mail gönderilemedi. Şifre veya portları kontrol edin. Detay: {e}")
        sys.exit(1)
else:
    print("Yeni risk haberi bulunamadı, mail atılmadı.")
