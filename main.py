import os
import re
import json
import time
import pandas as pd
import feedparser
import psycopg2
import smtplib
import hashlib
import unicodedata
from datetime import datetime, timezone
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import google.generativeai as genai

# --- AYARLAR VE GİZLİ ANAHTARLAR ---
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL") 
GONDEREN_MAIL = os.environ.get("GONDEREN_MAIL")
GONDEREN_SIFRE = os.environ.get("GONDEREN_SIFRE")
ALICI_MAILLER = ["muratfirar@hotmail.com.tr","muratfirar@isleasing.com.tr"]

if not all([GEMINI_API_KEY, SUPABASE_URL, GONDEREN_MAIL, GONDEREN_SIFRE]):
    print("KRİTİK HATA: Ortam değişkenleri (Secrets) eksik!")
    exit(1)

# Model 404 hatasını çözmek için geçerli model atandı[cite: 1]
GEMINI_MODEL = "gemini-3.5-flash-lite" 
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel(GEMINI_MODEL)

ANAHTAR_KELIMELER = ["konkordato", "iflas", "dava", "haciz", "grev", "ceza", "zarar", "temerrüt", "temerrüd", "soruşturma"][cite: 2]
NEGATIF_KELIMELER = ["bedava", "promosyon", "reklam", "sponsor", "indirim", "dizi"][cite: 2]

def normalize_metin(metin):
    metin = str(metin).casefold()
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(c for c in metin if not unicodedata.combining(c))[cite: 2]

def regex_eslesme(metin, ifade):
    metin_n = normalize_metin(metin)
    ifade_n = normalize_metin(ifade)
    # Sağ kelime sınırı yok, sondan eklemeli kelimeleri (temerrüde) yakalar[cite: 2]
    return bool(re.search(r"(?<!\w)" + re.escape(ifade_n), metin_n))[cite: 2]

def bulunanlar(metin, kelimeler):
    return [k for k in kelimeler if regex_eslesme(metin, k)][cite: 4]

def sha256_text(text):
    return hashlib.sha256((text or "").encode("utf-8", errors="ignore")).hexdigest()[cite: 4]

# --- SUPABASE POSTGRESQL KURULUMU ---
conn = psycopg2.connect(SUPABASE_URL)[cite: 4]
cursor = conn.cursor()

cursor.execute("""
    CREATE TABLE IF NOT EXISTS arama_gecmisi (
        arama_id SERIAL PRIMARY KEY,
        sorgu_zamani TEXT NOT NULL,
        firma_sayisi INTEGER,
        rss_sonucu INTEGER,
        aday_sayisi INTEGER,
        gemini_cagrisi INTEGER,
        hata_sayisi INTEGER
    )
""")[cite: 4]

cursor.execute("""
    CREATE TABLE IF NOT EXISTS haber_eslesmeleri (
        haber_id SERIAL PRIMARY KEY,
        firma_unvan TEXT NOT NULL,
        url TEXT UNIQUE NOT NULL,
        baslik TEXT,
        haber_tarihi TEXT,
        firma_eslesmeleri TEXT,
        risk_kelimeleri TEXT,
        negatif_eslesmeleri TEXT,
        kural_skoru INTEGER,
        gemini_sinifi TEXT,
        gemini_gerekce TEXT,
        gemini_guven REAL,
        nihai_durum TEXT,
        gemini_model TEXT,
        icerik_hash TEXT,
        ilk_gorulme TEXT
    )
""")[cite: 4]
conn.commit()

# --- EXCEL OKUMA ---
try:
    df_firmalar = pd.read_excel('firmalar.xlsx')
except Exception as e:
    print(f"Excel hatası: {e}")
    exit(1)

tum_sonuclar = []
rss_toplam = 0
gemini_cagrisi = 0
hatalar = 0
zaman = datetime.now(timezone.utc).isoformat()[cite: 4]

# --- GELİŞMİŞ TARAMA VE DEĞERLENDİRME ---
for index, row in df_firmalar.iterrows():
    firma_tam = str(row.get('firma_unvan', '')).strip()
    firma_kisa = str(row.get('haber_unvan', '')).strip()
    
    keyword_query = " OR ".join(ANAHTAR_KELIMELER)
    sorgu = f'"{firma_kisa}" AND ({keyword_query})'
    rss_url = f"https://news.google.com/rss/search?q={urllib.parse.quote(sorgu)}&hl=tr&gl=TR"
    
    try:
        feed = feedparser.parse(rss_url)
        rss_toplam += len(feed.entries)
        
        for entry in feed.entries:
            baslik = entry.title
            link = entry.link
            tarih = entry.published[5:16] if hasattr(entry, 'published') else "Bilinmeyen"
            
            # Kural 1: Negatif Eşleşme (Blacklist)
            negatif_eslesmeler = bulunanlar(baslik, NEGATIF_KELIMELER)[cite: 4]
            if negatif_eslesmeler:
                continue
                
            # Kural 2: Risk Kelimeleri ve Firma Teyidi
            risk_eslesmeleri = bulunanlar(baslik, ANAHTAR_KELIMELER)[cite: 4]
            firma_eslesmeleri = bulunanlar(baslik, [firma_kisa])[cite: 4]
            
            kural_skoru = (40 if firma_eslesmeleri else 0) + (30 if risk_eslesmeleri else 0)[cite: 2]
            icerik_hash = sha256_text(baslik)[cite: 4]
            
            # Supabase Tekrar Kontrolü (Önbellek)
            cursor.execute("SELECT gemini_sinifi, gemini_gerekce, gemini_guven, gemini_model FROM haber_eslesmeleri WHERE url=%s", (link,))[cite: 4]
            cached = cursor.fetchone()
            
            gemini_sinifi, gemini_gerekce, gemini_guven, model_kullanilan = "ON_FILTREDE_ELENDI", "", 0.0, ""
            nihai_durum = "ON_FILTREDE_ELENDI"
            
            should_call = bool(firma_eslesmeleri and risk_eslesmeleri)[cite: 2]
            
            if should_call:
                if cached and cached[0] is not None:
                    gemini_sinifi, gemini_gerekce, gemini_guven, model_kullanilan = cached
                    nihai_durum = "CACHE_YUKLENDI"
                else:
                    # Gemini JSON Prompt Formatı[cite: 4]
                    prompt = f"""
                    Sen Türkçe kurumsal haber taraması yapan bir analistsin.
                    Firma: {firma_tam}
                    Haber Başlığı: {baslik}
                    JSON dışında hiçbir şey döndürme ve şu alanları kullan: 
                    {{ "ilgili_firma": true, "olumsuz_olay": true, "olay_turu": "finansal|hukuki", "sinif": "OLUMLU_ADAY|ILGISIZ|BELIRSIZ", "guven": 0.0, "gerekce": "Kısa kanıt." }}
                    Kurallar: "OLUMLU_ADAY" yalnızca haber hedef firmayla ilgili VE somut olumsuz olay içeriyorsa seçilsin.
                    """
                    try:
                        resp = model.generate_content(
                            prompt,
                            generation_config={"temperature": 0.1, "response_mime_type": "application/json"}
                        )[cite: 4]
                        data = json.loads(resp.text.strip())
                        gemini_sinifi = str(data.get("sinif", "BELIRSIZ")).upper()[cite: 4]
                        gemini_guven = max(0.0, min(1.0, float(data.get("guven", 0.0))))[cite: 4]
                        gemini_gerekce = str(data.get("gerekce", ""))[:1000][cite: 4]
                        model_kullanilan = GEMINI_MODEL
                        gemini_cagrisi += 1
                        
                        if gemini_sinifi == "OLUMLU_ADAY" and data.get("ilgili_firma") and data.get("olumsuz_olay"):[cite: 4]
                            nihai_durum = "GEMINI_ADAY_INCELE"[cite: 4]
                        else:
                            nihai_durum = f"GEMINI_{gemini_sinifi}"[cite: 4]
                    except Exception as e:
                        gemini_sinifi = "API_HATASI"[cite: 4]
                        gemini_gerekce = str(e)[:200]
                        nihai_durum = "API_HATASI"[cite: 4]

                # Supabase Kayıt (Postgres Upsert)[cite: 4]
                cursor.execute("""
                    INSERT INTO haber_eslesmeleri (
                        firma_unvan, url, baslik, haber_tarihi, firma_eslesmeleri, risk_kelimeleri, negatif_eslesmeleri, 
                        kural_skoru, gemini_sinifi, gemini_gerekce, gemini_guven, nihai_durum, gemini_model, icerik_hash, ilk_gorulme
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (url) DO UPDATE SET 
                        kural_skoru=EXCLUDED.kural_skoru, 
                        gemini_sinifi=EXCLUDED.gemini_sinifi, gemini_gerekce=EXCLUDED.gemini_gerekce, 
                        gemini_guven=EXCLUDED.gemini_guven, nihai_durum=EXCLUDED.nihai_durum
                """, (
                    firma_tam, link, baslik, tarih, ", ".join(firma_eslesmeleri), ", ".join(risk_eslesmeleri), 
                    ", ".join(negatif_eslesmeler), kural_skoru, gemini_sinifi, gemini_gerekce, gemini_guven, 
                    nihai_durum, model_kullanilan, icerik_hash, zaman
                ))
                conn.commit()
                
                if nihai_durum == "GEMINI_ADAY_INCELE":[cite: 4]
                    tum_sonuclar.append({
                        "Firma": firma_tam, "Risk Faktörü": ", ".join(risk_eslesmeleri), "Skor": kural_skoru, 
                        "Güven": gemini_guven, "Gerekçe": gemini_gerekce, "Başlık": baslik, "Link": link
                    })
    except Exception as e:
        hatalar += 1
    time.sleep(1.5)

# Arama Geçmişi Log Kaydı[cite: 4]
cursor.execute("""
    INSERT INTO arama_gecmisi (sorgu_zamani, firma_sayisi, rss_sonucu, aday_sayisi, gemini_cagrisi, hata_sayisi) 
    VALUES (%s, %s, %s, %s, %s, %s)
""", (zaman, len(df_firmalar), rss_toplam, len(tum_sonuclar), gemini_cagrisi, hatalar))[cite: 4]
conn.commit()
cursor.close()
conn.close()

# --- E-POSTA BİLDİRİMİ ---
if tum_sonuclar:
    df_rapor = pd.DataFrame(tum_sonuclar)
    html_tablo = df_rapor.to_html(escape=False, index=False, formatters={'Link': lambda v: f'<a href="{v}">Haberi Oku</a>'})
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"⚠️ Erken Uyarı: {len(df_rapor)} Yeni Aday Bulundu"
    msg["From"] = GONDEREN_MAIL
    msg["To"] = ", ".join(ALICI_MAILLER)
    msg.attach(MIMEText(f"<h2>Günlük Risk Tarama Raporu</h2>{html_tablo}", "html"))
    
    try:
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(GONDEREN_MAIL, GONDEREN_SIFRE)
        server.sendmail(GONDEREN_MAIL, ALICI_MAILLER, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Mail hatası: {e}")
