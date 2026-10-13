import smtplib
import ssl # Bu satırı ekleyin
from email.mime.multipart import MIMEMultipart
import os
import re
import json
import time
import html
import hashlib
import unicodedata
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timezone
from urllib.parse import urlencode, urlparse

import pandas as pd
import requests
import feedparser
import trafilatura
from bs4 import BeautifulSoup
from google import genai
from google.genai import types
import psycopg2

# -------- 1. AYARLAR VE ÇEVRE DEĞİŞKENLERİ --------
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL")
GONDEREN_MAIL = os.environ.get("GONDEREN_MAIL")
GONDEREN_SIFRE = os.environ.get("GMAIL_APP_PASSWORD")

# E-postanın gideceği 2 adresi buraya tanımlayın
ALICI_MAILLER_STR = os.environ.get("ALICI_MAILLER", "") 
ALICI_MAILLER = [m.strip() for m in ALICI_MAILLER_STR.split(",") if m.strip()]

if not ALICI_MAILLER:
    print("KRİTİK HATA: ALICI_MAILLER çevre değişkeni bulunamadı!")
    exit(1)

#ALICI_MAILLER = ["muratfirar@hotmail.com.tr", "muratfirar@isleasing.com.tr"]

if not all([GEMINI_API_KEY, SUPABASE_URL, GONDEREN_MAIL, GONDEREN_SIFRE]):
    print("HATA: Çevre değişkenleri (Secrets) eksik!")
    exit(1)

client = genai.Client(api_key=GEMINI_API_KEY)

GEMINI_MODEL_ISTENEN = "gemini-3.5-flash-lite"
GEMINI_MODEL_YEDEK =   "gemini-3.1-flash-lite"
MODEL_YEDEK_KULLAN = True

MAX_HABER_FIRMA = 8
MAX_GEMINI_CAGRISI = 450
BEKLEME_RSS_SANIYE = 1.5
BEKLEME_SAYFA_SANIYE = 4.5
HABER_GOVDESI_INDIR = True
GOVDE_MAX_KARAKTER = 9000
RSS_ACIKLAMA_MAX_KARAKTER = 5000
EXCEL_DOSYASI = "Firmalar.xlsx"

# -------- 2. KELİME LİSTELERİ VE YARDIMCI FONKSİYONLAR --------
anahtar_kelimeler = [
    "iflas", "konkordato", "icra", "haciz", "dava", "soruşturma", "ceza", 
    "temerrüt", "temerrüd", "borç", "dolandırıcılık", "faaliyet durdurma", 
    "kayyum", "yeniden yapılandırma", "borçlarını ödeyemedi", "tedbir", 
    "tazminat", "usulsüzlük", "yaptırım", "lisans iptali", "ödeme güçlüğü"
]
genel_negatifler = ["bedava", "promosyon", "reklam", "sponsor", "indirim", "dizi"]

def rss_anahtar_ifadesi(kelimeler):
    ifadeler = [f'"{k.strip()}"' if " " in k else k.strip() for k in kelimeler if k.strip()]
    return " OR ".join(ifadeler)

def normalize_metin(metin):
    metin = str(metin).casefold()
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(c for c in metin if not unicodedata.combining(c))

def regex_eslesme(metin, ifade): 
    metin_n = normalize_metin(metin) 
    ifade_n = normalize_metin(ifade) 
    return bool(re.search(r"(?<!\w)" + re.escape(ifade_n), metin_n))

def bulunanlar(metin, kelimeler):
    return [k for k in kelimeler if regex_eslesme(metin, k)]

def sha256_text(text):
    return hashlib.sha256((text or "").encode("utf-8", errors="ignore")).hexdigest()

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; FirmaHaberResearch/1.0)"})

def haber_ara(sorgu, limit=8):
    rss_url = "https://news.google.com/rss/search?" + urlencode({"q": sorgu, "hl": "tr", "gl": "TR", "ceid": "TR:tr"})
    try:
        response = session.get(rss_url, timeout=25)
        response.raise_for_status()
        feed = feedparser.parse(response.content)
        rows = []
        for e in feed.entries[:limit]:
            summary_html = e.get("summary", "")
            summary = BeautifulSoup(summary_html, "html.parser").get_text(" ", strip=True)
            rows.append({
                "baslik": e.get("title", ""), "url": e.get("link", ""),
                "tarih": e.get("published", ""), "rss_aciklama": summary[:RSS_ACIKLAMA_MAX_KARAKTER]
            })
        return rows, None
    except Exception as exc:
        return [], str(exc)[:400]

def haber_govdesi_getir(url):
    if not url or urlparse(url).scheme not in ("http", "https"): return "", "GECERSIZ_URL"
    try:
        response = session.get(url, timeout=25)
        response.raise_for_status()
        if "html" not in response.headers.get("Content-Type", "").lower(): return "", "HTML_DEGIL"
        text = trafilatura.extract(response.text, url=url, include_comments=False, include_tables=False, favor_precision=True)
        if text and len(text.strip()) >= 100: return text.strip()[:GOVDE_MAX_KARAKTER], "BASARILI"
        return "", "METIN_CIKARILAMADI"
    except Exception as exc:
        return "", "HATA: " + str(exc)[:180]

# -------- 3. VERİTABANI BAĞLANTISI (POSTGRESQL) --------
conn = psycopg2.connect(SUPABASE_URL)
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
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS haber_eslesmeleri (
    haber_id SERIAL PRIMARY KEY,
    firma_unvan TEXT NOT NULL,
    url TEXT NOT NULL,
    baslik TEXT,
    haber_tarihi TEXT,
    rss_aciklama TEXT,
    govde TEXT,
    govde_durumu TEXT,
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
    ilk_gorulme TEXT,
    son_gorulme TEXT,
    son_arama_id INTEGER,
    UNIQUE(firma_unvan, url)
)
""")
conn.commit()

def cached_gemini_result(firma_unvan, url, content_hash):
    cursor.execute("""
        SELECT gemini_sinifi, gemini_gerekce, gemini_guven, gemini_model
        FROM haber_eslesmeleri
        WHERE firma_unvan=%s AND url=%s AND icerik_hash=%s AND gemini_sinifi IS NOT NULL
    """, (firma_unvan, url, content_hash))
    row = cursor.fetchone()
    if not row: return None
    return {"sinif": row[0], "gerekce": row[1], "guven": row[2], "model": row[3]}

def kural_on_degerlendirme(firma, haber):
    baslik, rss, govde = haber.get("baslik", ""), haber.get("rss_aciklama", ""), haber.get("govde", "")
    metin = baslik + "\n" + rss + "\n" + govde
    isimler = list(dict.fromkeys([firma["firma_unvan"], firma["haber_unvan"]] + firma["alternatif_unvanlar"]))
    
    firma_baslik = bulunanlar(baslik, isimler)
    firma_rss = bulunanlar(rss, isimler)
    firma_govde = bulunanlar(govde, isimler)
    
    risk_baslik = bulunanlar(baslik, anahtar_kelimeler)
    risk_rss = bulunanlar(rss, anahtar_kelimeler)
    risk_govde = bulunanlar(govde, anahtar_kelimeler)
    
    negatifler = list(dict.fromkeys(genel_negatifler + firma.get("ozel_negatifler", [])))
    
    skor = 0
    if firma_baslik: skor += 40
    elif firma_rss: skor += 25
    elif firma_govde: skor += 20

    if risk_baslik: skor += 30
    elif risk_rss: skor += 15
    elif risk_govde: skor += 10

    marti_jenerik = (firma["firma_unvan"] == "Martı İleri Teknoloji A.Ş." and bool(bulunanlar(baslik + " " + rss, ["kuş", "sahil", "plaj", "kanat", "göç"])) and not bool(firma_baslik or firma_rss))

    return {
        "firma_eslesmesi_var": bool(firma_baslik or firma_rss or firma_govde),
        "risk_eslesmesi_var": bool(risk_baslik or risk_rss or risk_govde),
        "firma_eslesmeleri": list(dict.fromkeys(firma_baslik + firma_rss + firma_govde)),
        "risk_kelimeleri": list(dict.fromkeys(risk_baslik + risk_rss + risk_govde)),
        "negatif_eslesmeleri": bulunanlar(metin, negatifler),
        "kural_skoru": min(skor, 100),
        "marti_jenerik": marti_jenerik
    }

def gemini_degerlendir(firma, haber, model_tercihi, model_yedek, fallback=True):
    body = (haber.get("govde") or "")[:GOVDE_MAX_KARAKTER]
    rss = (haber.get("rss_aciklama") or "")[:RSS_ACIKLAMA_MAX_KARAKTER]
    prompt = f"""
Sen Türkçe kurumsal haber taraması yapan bir analistsin. Yalnızca aşağıdaki haber verisine dayan.
Hedef firma tam unvanı: {firma['firma_unvan']}
Haberde kullanılan ad: {firma['haber_unvan']}
Alternatif adlar: {", ".join(firma['alternatif_unvanlar'])}
Başlık: {haber.get('baslik','')}
RSS açıklaması: {rss}
Gövde: {body if body else '[Çıkarılamadı]'}
JSON dışında hiçbir şey döndürme: {{"ilgili_firma": true, "olumsuz_olay": true, "olay_turu": "finansal|hukuki|belirsiz", "sinif": "OLUMLU_ADAY|ILGISIZ|BELIRSIZ", "guven": 0.0, "gerekce": "Kısa kanıt."}}
Kurallar: "OLUMLU_ADAY" yalnızca haber hedef firmayla ilgili VE somut olumsuz olay içeriyorsa seçilsin.
"""
    models = [model_tercihi] + ([model_yedek] if fallback and model_yedek != model_tercihi else [])
    last_error = None
    for m in models:
        try:
            response = client.models.generate_content(model=m, contents=prompt, config=types.GenerateContentConfig(temperature=0.1, response_mime_type="application/json"))
            data = json.loads(response.text or "{}")
            return {
                "ilgili_firma": bool(data.get("ilgili_firma", False)), "olumsuz_olay": bool(data.get("olumsuz_olay", False)),
                "olay_turu": str(data.get("olay_turu", "belirsiz")), "sinif": str(data.get("sinif", "BELIRSIZ")).upper(),
                "guven": max(0.0, min(1.0, float(data.get("guven", 0.0)))), "gerekce": str(data.get("gerekce", ""))[:1000],
                "model": m, "hata": ""
            }
        except Exception as exc:
            last_error = str(exc)[:500]
            continue
    return {"ilgili_firma": False, "olumsuz_olay": False, "olay_turu": "belirsiz", "sinif": "API_HATASI", "guven": 0.0, "gerekce": last_error or "Hata", "model": "", "hata": last_error or "Hata"}

# -------- 4. ANA İŞ AKIŞI VE TARAMA --------
def excel_hucre_cozumle(hucre_verisi):
    """Excel hücresindeki veriyi (JSON veya virgülle ayrılmış metin) güvenle listeye çevirir."""
    if pd.isna(hucre_verisi) or str(hucre_verisi).strip() == "":
        return []
    
    veri_metni = str(hucre_verisi).strip()
    # Eğer daha önceki gibi JSON formatında yazılmışsa [ "isim1", "isim2" ]
    if veri_metni.startswith("["):
        try:
            return json.loads(veri_metni)
        except json.JSONDecodeError:
            pass
    
    # Normal virgülle ayrılmış metin olarak yazılmışsa (Örn: THY, Türk Hava Yolları)
    return [isim.strip() for isim in veri_metni.split(",") if isim.strip()]

try:
    df_excel = pd.read_excel(EXCEL_DOSYASI)
    firmalar = []
    for _, row in df_excel.iterrows():
        haber_unvan_str = str(row['haber_unvan']).strip()
        
        # Alternatif unvanlar sütunu varsa ve doluysa al, yoksa sadece haber_unvan'ı kullan
        alt_unvanlar = excel_hucre_cozumle(row.get('alternatif_unvanlar'))
        if not alt_unvanlar:
            alt_unvanlar = [haber_unvan_str]
            
        firmalar.append({
            "firma_unvan": str(row['firma_unvan']).strip(),
            "haber_unvan": haber_unvan_str,
            "alternatif_unvanlar": alt_unvanlar,
            "ozel_negatifler": excel_hucre_cozumle(row.get('ozel_negatifler'))
        })
    print(f"Excel'den {len(firmalar)} firma başarıyla okundu.")
except Exception as e:
    print(f"Excel okunurken hata oluştu: {e}")
    exit(1)

tum_sonuclar = []
gemini_cagrisi_sayisi, rss_toplam, arama_hatalari = 0, 0, []
zaman = datetime.now(timezone.utc).isoformat()

for idx, firma in enumerate(firmalar, start=1):
    anahtar_ifadesi = rss_anahtar_ifadesi(anahtar_kelimeler)
    sorgu = f'"{firma["haber_unvan"]}" ({anahtar_ifadesi})'
    
    haberler, hata = haber_ara(sorgu, MAX_HABER_FIRMA)
    if hata: 
        arama_hatalari.append({"firma": firma["haber_unvan"], "hata": hata})
        continue
    
    rss_toplam += len(haberler)
    for haber in haberler:
        if HABER_GOVDESI_INDIR:
            govde, govde_durumu = haber_govdesi_getir(haber["url"])
            haber["govde"], haber["govde_durumu"] = govde, govde_durumu
            time.sleep(BEKLEME_SAYFA_SANIYE)
        else:
            haber["govde"], haber["govde_durumu"] = "", "ATLANDI"

        kural = kural_on_degerlendirme(firma, haber)
        content_hash = sha256_text(haber.get("baslik", "") + "\n" + haber.get("rss_aciklama", "") + "\n" + haber.get("govde", ""))
        
        gemini = {"sinif": "ON_FILTREDE_ELENDI", "gerekce": "", "guven": 0.0, "model": "", "hata": ""}
        should_call = kural["firma_eslesmesi_var"] and kural["risk_eslesmesi_var"] and not kural["marti_jenerik"]
        cached = cached_gemini_result(firma["firma_unvan"], haber["url"], content_hash)

        if should_call and cached:
            gemini.update({"sinif": cached["sinif"], "gerekce": cached["gerekce"] or "", "guven": float(cached["guven"] or 0.0), "model": cached["model"] or "", "cache": True})
        elif should_call and gemini_cagrisi_sayisi < MAX_GEMINI_CAGRISI:
            gemini = gemini_degerlendir(firma, haber, GEMINI_MODEL_ISTENEN, GEMINI_MODEL_YEDEK, MODEL_YEDEK_KULLAN)
            gemini_cagrisi_sayisi += 1; gemini["cache"] = False
            time.sleep(4.5)
        elif should_call:
            gemini.update({"sinif": "GEMINI_LIMITI", "gerekce": "Limit aşıldı."})

        if kural["marti_jenerik"]: durum = "ELENDI_JENERIK_BAGLAM"
        elif gemini["sinif"] == "OLUMLU_ADAY" and gemini.get("ilgili_firma", True) and gemini.get("olumsuz_olay", True): durum = "GEMINI_ADAY_INCELE"
        elif gemini["sinif"] == "ILGISIZ": durum = "GEMINI_ILGISIZ"
        elif gemini["sinif"] == "BELIRSIZ": durum = "GEMINI_BELIRSIZ"
        elif gemini["sinif"] in ["API_HATASI", "GEMINI_LIMITI"]: durum = gemini["sinif"]
        elif not should_call: durum = "ON_FILTREDE_ELENDI"
        else: durum = "INCELEME_GEREKLI"

        tum_sonuclar.append({
            "firma_unvan": firma["firma_unvan"], "baslik": haber.get("baslik", ""), "url": haber.get("url", ""),
            "tarih": haber.get("tarih", ""), "rss_aciklama": haber.get("rss_aciklama", ""), "govde": haber.get("govde", ""),
            "govde_durumu": haber.get("govde_durumu", ""), "firma_eslesmeleri": ", ".join(kural["firma_eslesmeleri"]),
            "risk_kelimeleri": ", ".join(kural["risk_kelimeleri"]), "negatif_eslesmeleri": ", ".join(kural["negatif_eslesmeleri"]),
            "kural_skoru": kural["kural_skoru"], "gemini_sinifi": gemini["sinif"], "gemini_gerekce": gemini["gerekce"],
            "gemini_guven": gemini["guven"], "gemini_model": gemini.get("model", ""), "nihai_durum": durum, "icerik_hash": content_hash
        })
    time.sleep(BEKLEME_RSS_SANIYE)

# -------- 5. VERİTABANINA YAZMA --------
aday_sayisi = sum(1 for r in tum_sonuclar if r["nihai_durum"] == "GEMINI_ADAY_INCELE")

cursor.execute("""
INSERT INTO arama_gecmisi (sorgu_zamani, firma_sayisi, rss_sonucu, aday_sayisi, gemini_cagrisi, hata_sayisi)
VALUES (%s, %s, %s, %s, %s, %s) RETURNING arama_id
""", (zaman, len(firmalar), rss_toplam, aday_sayisi, gemini_cagrisi_sayisi, len(arama_hatalari)))
arama_id = cursor.fetchone()[0]

for r in tum_sonuclar:
    cursor.execute("""
    INSERT INTO haber_eslesmeleri (
        firma_unvan, url, baslik, haber_tarihi, rss_aciklama, govde, govde_durumu, firma_eslesmeleri, 
        risk_kelimeleri, negatif_eslesmeleri, kural_skoru, gemini_sinifi, gemini_gerekce, gemini_guven, 
        nihai_durum, gemini_model, icerik_hash, ilk_gorulme, son_gorulme, son_arama_id
    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    ON CONFLICT (firma_unvan, url) DO UPDATE SET
        baslik=EXCLUDED.baslik, haber_tarihi=EXCLUDED.haber_tarihi, rss_aciklama=EXCLUDED.rss_aciklama,
        govde=CASE WHEN EXCLUDED.govde IS NOT NULL AND EXCLUDED.govde != '' THEN EXCLUDED.govde ELSE haber_eslesmeleri.govde END,
        govde_durumu=EXCLUDED.govde_durumu, firma_eslesmeleri=EXCLUDED.firma_eslesmeleri, risk_kelimeleri=EXCLUDED.risk_kelimeleri,
        negatif_eslesmeleri=EXCLUDED.negatif_eslesmeleri, kural_skoru=EXCLUDED.kural_skoru, gemini_sinifi=EXCLUDED.gemini_sinifi,
        gemini_gerekce=EXCLUDED.gemini_gerekce, gemini_guven=EXCLUDED.gemini_guven, nihai_durum=EXCLUDED.nihai_durum,
        gemini_model=EXCLUDED.gemini_model, icerik_hash=EXCLUDED.icerik_hash, son_gorulme=EXCLUDED.son_gorulme, son_arama_id=EXCLUDED.son_arama_id
    """, (
        r["firma_unvan"], r["url"], r["baslik"], r["tarih"], r["rss_aciklama"], r["govde"], r["govde_durumu"],
        r["firma_eslesmeleri"], r["risk_kelimeleri"], r["negatif_eslesmeleri"], int(r["kural_skoru"]),
        r["gemini_sinifi"], r["gemini_gerekce"], float(r["gemini_guven"]), r["nihai_durum"], r["gemini_model"],
        r["icerik_hash"], zaman, zaman, arama_id
    ))
conn.commit()
cursor.close()
conn.close()

# -------- 6. E-POSTA BİLDİRİMİ --------
aday_haberler = [r for r in tum_sonuclar if r["nihai_durum"] == "GEMINI_ADAY_INCELE"]

msg = MIMEMultipart("alternative")
msg["From"] = GONDEREN_MAIL
#msg["To"] = ", ".join(ALICI_MAILLER)
msg["To"] = "undisclosed-recipients

if aday_haberler:
    df_rapor = pd.DataFrame(aday_haberler)[["firma_unvan", "baslik", "tarih", "risk_kelimeleri", "gemini_gerekce", "url"]]
    df_rapor["url"] = df_rapor["url"].apply(lambda x: f'<a href="{x}">Habere Git</a>')
    html_tablo = df_rapor.to_html(escape=False, index=False)
    
    msg["Subject"] = f"⚠️ Erken Uyarı: {len(aday_haberler)} Yeni Risk Haberi Bulundu"
    html_icerik = f"<h2>Günlük Risk Tarama Raporu</h2><p>Aşağıdaki firmalar için risk teşkil edebilecek haberler tespit edildi:</p>{html_tablo}"
else:
    msg["Subject"] = "ℹ️ Günlük Tarama Raporu: Yeni Risk Bulunmadı"
    html_icerik = "<h2>Günlük Risk Tarama Raporu</h2><p>Bugünkü internet taramasında belirlenen kriterlere (iflas, konkordato vb.) takılan yeni bir risk haberi tespit edilmemiştir.</p>"

msg.attach(MIMEText(html_icerik, "html"))

try:
   
    # Port 587 kullanılarak bağlantı başlatılıyor (30 saniye zaman aşımı sınırı ile)
    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
    server.ehlo() # Sunucuya merhaba de
    server.set_debuglevel(1)
    server.starttls()
    server.ehlo() # Tünel içinden tekrar merhaba de
    
    server.login(GONDEREN_MAIL, GONDEREN_SIFRE)
    server.sendmail(GONDEREN_MAIL, ALICI_MAILLER, msg.as_string())
    server.quit()
    print("E-posta başarıyla gönderildi.")
except Exception as e:
    print(f"E-posta gönderim hatası: {e}")
