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

# --- 1. AYARLAR VE ÇEVRE DEĞİŞKENLERİ ---
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL") 
GONDEREN_MAIL = os.environ.get("GONDEREN_MAIL")
GONDEREN_SIFRE = os.environ.get("GONDEREN_SIFRE")
ALICI_MAILLER = ["muratfirar@hotmail.com","muratfirar@isleasing.com.tr"] # Raporun gideceği mail adres(ler)ini buraya yaz

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-3.5-flash-latest')

ANAHTAR_KELIMELER = ["konkordato", "iflas", "dava", "haciz", "grev", "ceza", "zarar", "temerrüt"]
NEGATIF_KELIMELER = {"Genel": ["bedava", "promosyon", "reklam", "sponsor", "indirim", "dizi"]}

def tr_lower(text):
    if not isinstance(text, str): return ""
    return text.replace("&quot;", '"').replace("&#39;", "'").replace("&nbsp;", " ").replace("İ", "i").replace("I", "ı").lower()

# --- 2. SUPABASE (POSTGRESQL) BAĞLANTISI VE TABLO ---
conn = psycopg2.connect(SUPABASE_URL)
cursor = conn.cursor()

# PostgreSQL formatında tablo oluşturma (SQLite'dan farklı olarak SERIAL kullanılır)
cursor.execute('''
    CREATE TABLE IF NOT EXISTS risk_haberleri (
        id SERIAL PRIMARY KEY,
        firma TEXT, 
        kaynak TEXT, 
        tarih TEXT, 
        baslik TEXT, 
        link TEXT UNIQUE, 
        llm_ozet TEXT
    )
''')
conn.commit()

# --- 3. EXCEL OKUMA VE HABER TARAMA ---
try:
    df_firmalar = pd.read_excel('firmalar.xlsx')
except Exception as e:
    print(f"Excel hatası: {e}")
    exit()

raporlanacak_haberler = []
toplam_firma = len(df_firmalar)

for index, row in df_firmalar.iterrows():
    firma_tam_ad = str(row['firma_unvan']).strip()
    haber_unvan = str(row['haber_unvan']).strip()
    
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
            
            # Negatif Kelime Filtresi
            if any(negatif in baslik_lower for negatif in firma_negatifleri):
                continue
                
            eslesenler = [kw for kw in ANAHTAR_KELIMELER if re.search(rf"(?:^|\W){kw}", baslik_lower)]
            
            if eslesenler:
                # LLM'e göndermeden hemen önce Supabase'de link kontrolü (Maliyet tasarrufu)
                cursor.execute("SELECT id FROM risk_haberleri WHERE link=%s", (link,))
                if cursor.fetchone() is not None:
                    continue # Haber Supabase'de var, pas geç
                
                # Sadece yeni haberler Gemini'a gider
                prompt = f"""
                Sen bir kurumsal kredi risk analistisin. 
                Aşağıdaki haber metni '{firma_tam_ad}' ile doğrudan ilgili mi ve finansal/hukuki bir risk içeriyor mu? 
                İlgiliyse 2 cümlelik bir yönetici özeti yaz. Önemsiz bir bahsedilmeyse veya ilgisizse sadece 'RED' yaz.
                Haber: {baslik}
                """
                try:
                    response = model.generate_content(prompt)
                    llm_yaniti = response.text.strip()
                except Exception as e:
                    llm_yaniti = "RED"
                
                # PostgreSQL formatında veritabanına kayıt (ON CONFLICT DO NOTHING)[cite: 1]
                cursor.execute('''
                    INSERT INTO risk_haberleri (firma, kaynak, tarih, baslik, link, llm_ozet)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (link) DO NOTHING
                ''', (firma_tam_ad, kaynak_adi, yayim_tarihi, baslik, link, llm_yaniti))
                conn.commit()
                
                # E-posta listesine sadece "RED" olmayanları ekle[cite: 1]
                if "RED" not in llm_yaniti.upper():
                    raporlanacak_haberler.append({
                        "Firma": firma_tam_ad, "Başlık": baslik, "Özet": llm_yaniti, "Bağlantı": link
                    })
                    
    except Exception as e:
        print(f"[{haber_unvan}] Hata: {e}")
        
    time.sleep(2) # Google Rate-Limit Koruması

cursor.close()
conn.close()

# --- 4. MAİL GÖNDERİMİ ---
if raporlanacak_haberler:
    df_sonuclar = pd.DataFrame(raporlanacak_haberler)
    def make_clickable(val):
        return f'<a href="{val}">Haberi Oku</a>'
    
    html_tablo = df_sonuclar.to_html(escape=False, index=False, formatters={'Bağlantı': make_clickable})
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
        print(f"Mail hatası: {e}")
else:
    print("Yeni risk haberi bulunamadı, mail atılmadı.")
