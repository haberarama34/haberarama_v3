import os
import re
import json
import time
import html
import hashlib
import unicodedata
import smtplib
import ssl
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

# E-postanın gideceği adresler
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
ANALIZ_SURUMU = "2.0"

MAX_HABER_FIRMA = 8
GRUP_BOYUTU = 5          # 2900 firmayı 5'erli gruplar halinde tarar (580 RSS sorgusu)
MAX_HABER_GRUP = 20      # Grup sorgusu başına çekilecek maksimum haber
MAX_GEMINI_CAGRISI = 450
BEKLEME_RSS_SANIYE = 1.5
BEKLEME_SAYFA_SANIYE = 4.5
HABER_GOVDESI_INDIR = True
GOVDE_MAX_KARAKTER = 9000
RSS_ACIKLAMA_MAX_KARAKTER = 5000
EXCEL_DOSYASI = "Firmalar_ikinci.xlsx"

# -------- 2. KELİME LİSTELERİ VE YARDIMCI FONKSİYONLAR --------
anahtar_kelimeler = [
    "iflas", "konkordato", "icra", "haciz", "dava", "soruşturma", "ceza", 
    "temerrüt", "temerrüd", "borç", "dolandırıcılık", "faaliyet durdurma", 
    "kayyum", "yeniden yapılandırma", "borçlarını ödeyemedi", "tedbir", 
    "tazminat", "usulsüzlük", "yaptırım", "lisans iptali", "ödeme güçlüğü"
]
genel_negatifler = ["bedava", "promosyon", "reklam", "sponsor", "indirim", "dizi"]

def otomatik_kisa_ad(tam_unvan):
    """Excel'de haber_unvan boş bırakıldığında resmi unvandan otomatik kısa marka adı türetir."""
    temiz = str(tam_unvan).upper()
    # Noktalama işaretlerini boşluğa çevir (Örn: A.B -> A B, ÖZ-SAN -> ÖZ SAN)
    temiz_kelime = re.sub(r"[^\w\s]", " ", temiz)
    
    # Resmi şirket uzantıları ve genel faaliyet kelimeleri
    ekler = {
        "ANONİM", "ANONIM", "ŞİRKETİ", "SIRKETI", "LİMİTED", "LIMITED", "SANAYİ", "SANAYI",
        "TİCARET", "TICARET", "VE", "A", "Ş", "S", "LTD", "ŞTİ", "STI", "SAN", "TİC", "TIC",
        "İNŞAAT", "INSAAT", "ELEKTRİK", "ELEKTRIK", "NAKLİYE", "NAKLIYE", "HURDACILIK",
        "DAYANIKLI", "TÜKETİM", "TUKETIM", "MALLARI", "DIŞ", "DIS", "YEDEK", "PARÇA", "PARCA",
        "İŞLEME", "ISLEME", "İTHALAT", "ITHALAT", "İHRACAT", "IHRACAT", "PAZARLAMA",
        "TAŞIMACILIK", "TASIMACILIK", "LOJİSTİK", "LOJISTIK", "HİZMETLERİ", "HIZMETLERI",
        "TURİZM", "TURIZM", "GIDA", "TEKSTİL", "TEKSTIL", "OTOMOTİV", "OTOMOTIV", "OTO",
        "METAL", "ÇELİK", "CELIK", "ALÜMİNYUM", "ALUMINYUM", "PLASTİK", "PLASTIK", "KİMYA",
        "KIMYA", "ENERJİ", "ENERJI", "TARIM", "HAYVANCILIK", "ÜRÜNLERİ", "URUNLERI", "ÜRETİM", "URETIM"
    }
    
    orijinal_kelimeler = str(tam_unvan).split()
    if not orijinal_kelimeler:
        return str(tam_unvan).strip()
        
    secilenler = []
    for k in orijinal_kelimeler:
        k_norm = re.sub(r"[^\w]", "", k.upper())
        if not k_norm:
            continue
        # İlk kelimeyi her zaman koru, sonraki kelimelerden en fazla 2 tanesini (ek değilse) al
        if len(secilenler) == 0:
            secilenler.append(k)
        elif k_norm not in ekler and len(secilenler) < 2:
            secilenler.append(k)
        else:
            # İkinci kelime bir sektör adıysa (Örn: Beyoğlu Çikolata, Arkat Alüminyum, A.B Gıda) onu da dahil et
            if len(secilenler) == 1 and k_norm in {"ÇİKOLATA", "GIDA", "ALÜMİNYUM", "ÇELİK", "VİNÇ", "OTO", "TEKSTİL", "HOLDİNG"}:
                secilenler.append(k)
            break
            
    kisa_ad = " ".join(secilenler)
    return kisa_ad.strip()

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
    if not ifade_n:
        return False
    return bool(re.search(r"(?<!\w)" + re.escape(ifade_n), metin_n))

def bulunanlar(metin, kelimeler):
    return [k for k in kelimeler if k and regex_eslesme(metin, k)]

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

# Yeni LLM analiz alanları; mevcut verileri silmeden tabloya eklenir.
cursor.execute("""
ALTER TABLE haber_eslesmeleri
 ADD COLUMN IF NOT EXISTS firma_eslesme_durumu TEXT,
 ADD COLUMN IF NOT EXISTS haberde_gecen_firma TEXT,
 ADD COLUMN IF NOT EXISTS firma_eslesme_kaniti TEXT,
 ADD COLUMN IF NOT EXISTS firma_eslesme_gerekcesi TEXT,
 ADD COLUMN IF NOT EXISTS firma_eslesme_guveni REAL,
 ADD COLUMN IF NOT EXISTS olay_turu TEXT,
 ADD COLUMN IF NOT EXISTS olay_ozeti TEXT,
 ADD COLUMN IF NOT EXISTS etkilenen_taraf TEXT,
 ADD COLUMN IF NOT EXISTS olayin_durumu TEXT,
 ADD COLUMN IF NOT EXISTS ciddiyet TEXT,
 ADD COLUMN IF NOT EXISTS hedef_firmaya_etki TEXT,
 ADD COLUMN IF NOT EXISTS olay_gerekcesi TEXT,
 ADD COLUMN IF NOT EXISTS olay_guveni REAL,
 ADD COLUMN IF NOT EXISTS analiz_surumu TEXT
""")
conn.commit()

def cached_gemini_result(firma_unvan, url, content_hash=None):
    cols = """gemini_sinifi, gemini_gerekce, gemini_guven, gemini_model, govde, govde_durumu, icerik_hash,
             firma_eslesme_durumu, haberde_gecen_firma, firma_eslesme_kaniti, firma_eslesme_gerekcesi, firma_eslesme_guveni,
             olay_turu, olay_ozeti, etkilenen_taraf, olayin_durumu, ciddiyet, hedef_firmaya_etki, olay_gerekcesi, olay_guveni, analiz_surumu"""
    if content_hash:
        cursor.execute(f"SELECT {cols} FROM haber_eslesmeleri WHERE firma_unvan=%s AND url=%s AND icerik_hash=%s AND gemini_sinifi IS NOT NULL AND analiz_surumu=%s", (firma_unvan,url,content_hash,ANALIZ_SURUMU))
    else:
        cursor.execute(f"SELECT {cols} FROM haber_eslesmeleri WHERE firma_unvan=%s AND url=%s AND gemini_sinifi IS NOT NULL AND analiz_surumu=%s ORDER BY son_gorulme DESC NULLS LAST LIMIT 1", (firma_unvan,url,ANALIZ_SURUMU))
    row=cursor.fetchone()
    if not row: return None
    return {
      "sinif":row[0],"gerekce":row[1] or "","guven":float(row[2] or 0),"model":row[3] or "","govde":row[4] or "","govde_durumu":row[5] or "","icerik_hash":row[6] or "",
      "firma_eslesme_durumu":row[7] or "DOGRULANAMADI","haberde_gecen_firma":row[8] or "","firma_eslesme_kaniti":row[9] or "","firma_eslesme_gerekcesi":row[10] or "","firma_eslesme_guveni":float(row[11] or 0),
      "olay_turu":row[12] or "BELIRSIZ","olay_ozeti":row[13] or "","etkilenen_taraf":row[14] or "","olayin_durumu":row[15] or "BELIRSIZ","ciddiyet":row[16] or "BELIRSIZ","hedef_firmaya_etki":row[17] or "BELIRLENEMEDI","olay_gerekcesi":row[18] or "","olay_guveni":float(row[19] or 0),"analiz_surumu":row[20] or ""
    }

def kural_on_degerlendirme(firma, haber):
    baslik, rss, govde = haber.get("baslik", ""), haber.get("rss_aciklama", ""), haber.get("govde", "")
    metin = baslik + "\n" + rss + "\n" + govde
    isimler = [i for i in dict.fromkeys([firma["firma_unvan"], firma["haber_unvan"]] + firma["alternatif_unvanlar"]) if i]
    
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

def _guven(v):
    try: return max(0.0,min(1.0,float(v)))
    except (TypeError,ValueError): return 0.0

def _metin(v, n=1500):
    if isinstance(v,list): return "; ".join(str(x).strip() for x in v if str(x).strip())[:n]
    return str(v or "")[:n]

def gemini_degerlendir(firma, haber, model_tercihi, model_yedek, fallback=True):
    body=(haber.get("govde") or "")[:GOVDE_MAX_KARAKTER]
    rss=(haber.get("rss_aciklama") or "")[:RSS_ACIKLAMA_MAX_KARAKTER]
    prompt=f"""
Sen Türkçe kurumsal haber taraması yapan bir analistsin. İKİ AYRI soruyu birbirine karıştırma: (1) Haberdeki şirket hedef firma mı? (2) Haberdeki olumsuz olay nedir ve hedef firmayı etkiliyor mu? Yalnızca verilen metne dayan; bilgi uydurma.

HEDEF FİRMA
Tam unvan: {firma['firma_unvan']}
Haberde kullanılan ad: {firma['haber_unvan']}
Alternatif adlar: {", ".join(firma['alternatif_unvanlar'])}

HABER
Başlık: {haber.get('baslik','')}
RSS açıklaması: {rss}
Gövde: {body if body else '[Çıkarılamadı]'}

FİRMA EŞLEŞMESİ KURALLARI
- Sadece kelime/harf/kısaltma benzerliği eşleşme değildir.
- Siyasi parti, kişi, kamu kurumu, marka veya başka şirket adının benzemesi kanıt değildir.
- Özellikle hedef unvanda bulunan bir kısaltmanın siyasi/kişisel kullanımını hedef şirket sayma.
- Haberdeki şirketin tam adı, unvanı, faaliyet alanı veya hedef firmayla açık bağlantısı varsa kanıt olarak kullan.
- Yeterli kanıt yoksa DOGRULANAMADI de.
Firma durumu: ESLESIYOR | ESLESMIYOR | DOGRULANAMADI | BELIRSIZ.

OLAY ANALİZİ
Olumsuz olay varsa ne olduğunu, etkilenen tarafı, olayın durumunu (KESINLESMIS/IDDİA/SORUSTURMA/KARAR/BELIRSIZ), türünü, ciddiyetini ve hedef firmaya etkisini ayrı değerlendir. Hedef firmaya doğrudan etkiyi varsayma.

NİHAİ SINIF
OLUMLU_ADAY yalnızca firma ESLESIYOR ve hedef firmayı ilgilendiren somut olumsuz olay varsa. Firma ESLESMIYOR ise normalde ILGISIZ. DOGRULANAMADI ise gerekirse INCELEME_GEREKLI.

Yalnızca şu JSON'u döndür:
{{
 "firma_eslesmesi": {{"durum":"ESLESIYOR|ESLESMIYOR|DOGRULANAMADI|BELIRSIZ","haberde_gecen_firma":"","hedef_firmayla_baglanti":"","eslesme_kaniti":[],"guven":0.0}},
 "olay_analizi": {{"olumsuz_olay_var":false,"olay_turu":"FINANSAL|HUKUKI|OPERASYONEL|ITIBARI|DUZENLEYICI|DIGER|BELIRSIZ","olay":"","etkilenen_taraf":"","olayin_durumu":"KESINLESMIS|IDDİA|SORUSTURMA|KARAR|BELIRSIZ","ciddiyet":"DUSUK|ORTA|YUKSEK|KRITIK|BELIRSIZ","hedef_firmaya_etki":"DOGRUDAN|DOLAYLI|ETKI_YOK|BELIRLENEMEDI","gerekce":"","guven":0.0}},
 "nihai_degerlendirme": {{"sinif":"OLUMLU_ADAY|ILGISIZ|BELIRSIZ|INCELEME_GEREKLI","inceleme_onceligi":"DUSUK|ORTA|YUKSEK|ACIL","ozet":"","guven":0.0}}
}}
"""
    models=[model_tercihi]+([model_yedek] if fallback and model_yedek!=model_tercihi else [])
    last_error=None
    for m in models:
        try:
            response=client.models.generate_content(model=m,contents=prompt,config=types.GenerateContentConfig(temperature=0.1,response_mime_type="application/json"))
            data=json.loads((response.text or "{}").strip().removeprefix("```json").removesuffix("```").strip())
            fe=data.get("firma_eslesmesi") or {}; oa=data.get("olay_analizi") or {}; nd=data.get("nihai_degerlendirme") or {}
            fd=str(fe.get("durum","DOGRULANAMADI")).upper().strip()
            if fd not in {"ESLESIYOR","ESLESMIYOR","DOGRULANAMADI","BELIRSIZ"}: fd="BELIRSIZ"
            sinif=str(nd.get("sinif","BELIRSIZ")).upper().strip()
            if sinif not in {"OLUMLU_ADAY","ILGISIZ","BELIRSIZ","INCELEME_GEREKLI"}: sinif="BELIRSIZ"
            if fd=="ESLESMIYOR" and sinif=="OLUMLU_ADAY": sinif="ILGISIZ"
            if fd=="DOGRULANAMADI" and sinif=="OLUMLU_ADAY": sinif="INCELEME_GEREKLI"
            return {
              "ilgili_firma":fd=="ESLESIYOR","olumsuz_olay":bool(oa.get("olumsuz_olay_var",False)),"olay_turu":str(oa.get("olay_turu","BELIRSIZ")).upper(),"sinif":sinif,"guven":_guven(nd.get("guven",0)),"gerekce":str(nd.get("ozet","") or "")[:1500],"model":m,"hata":"","cache":False,
              "firma_eslesme_durumu":fd,"haberde_gecen_firma":str(fe.get("haberde_gecen_firma","") or "")[:500],"firma_eslesme_kaniti":_metin(fe.get("eslesme_kaniti",[]),1200),"firma_eslesme_gerekcesi":str(fe.get("hedef_firmayla_baglanti","") or "")[:1500],"firma_eslesme_guveni":_guven(fe.get("guven",0)),
              "olay_ozeti":str(oa.get("olay","") or "")[:1500],"etkilenen_taraf":str(oa.get("etkilenen_taraf","") or "")[:500],"olayin_durumu":str(oa.get("olayin_durumu","BELIRSIZ")).upper(),"ciddiyet":str(oa.get("ciddiyet","BELIRSIZ")).upper(),"hedef_firmaya_etki":str(oa.get("hedef_firmaya_etki","BELIRLENEMEDI")).upper(),"olay_gerekcesi":str(oa.get("gerekce","") or "")[:1500],"olay_guveni":_guven(oa.get("guven",0)),"inceleme_onceligi":str(nd.get("inceleme_onceligi","DUSUK")).upper()
            }
        except Exception as exc: last_error=str(exc)[:500]
    return {"ilgili_firma":False,"olumsuz_olay":False,"olay_turu":"BELIRSIZ","sinif":"API_HATASI","guven":0.0,"gerekce":last_error or "Hata","model":"","hata":last_error or "Hata","cache":False,"firma_eslesme_durumu":"BELIRSIZ","haberde_gecen_firma":"","firma_eslesme_kaniti":"","firma_eslesme_gerekcesi":"","firma_eslesme_guveni":0.0,"olay_ozeti":"","etkilenen_taraf":"","olayin_durumu":"BELIRSIZ","ciddiyet":"BELIRSIZ","hedef_firmaya_etki":"BELIRLENEMEDI","olay_gerekcesi":"","olay_guveni":0.0,"inceleme_onceligi":"DUSUK"}

# -------- 4. ANA İŞ AKIŞI VE TARAMA --------

def excel_hucre_cozumle(hucre_verisi):
    """Excel hücresindeki veriyi (JSON veya virgülle ayrılmış metin) güvenle listeye çevirir."""
    if pd.isna(hucre_verisi) or str(hucre_verisi).strip() == "" or str(hucre_verisi).strip().lower() == "nan":
        return []
    
    veri_metni = str(hucre_verisi).strip()
    if veri_metni.startswith("["):
        try:
            return json.loads(veri_metni)
        except json.JSONDecodeError:
            pass
    
    return [isim.strip() for isim in veri_metni.split(",") if isim.strip()]

try:
    df_excel = pd.read_excel(EXCEL_DOSYASI)
    firmalar = []
    for _, row in df_excel.iterrows():
        if pd.isna(row.get('firma_unvan')) or str(row.get('firma_unvan')).strip() == "":
            continue
            
        firma_unvan_str = str(row['firma_unvan']).strip()
        
        # EĞER haber_unvan BOŞSA OTOMATİK KISA AD ÜRET, DOLUYSA MEVCUDU KULLAN
        ham_haber_unvan = row.get('haber_unvan')
        if pd.notna(ham_haber_unvan) and str(ham_haber_unvan).strip() != "" and str(ham_haber_unvan).strip().lower() != "nan":
            haber_unvan_str = str(ham_haber_unvan).strip()
        else:
            haber_unvan_str = otomatik_kisa_ad(firma_unvan_str)
        
        # Alternatif unvanlar sütunu varsa ve doluysa al, yoksa haber_unvan'ı kullan
        alt_unvanlar = excel_hucre_cozumle(row.get('alternatif_unvanlar'))
        if not alt_unvanlar:
            alt_unvanlar = [haber_unvan_str]
            
        firmalar.append({
            "firma_unvan": firma_unvan_str,
            "haber_unvan": haber_unvan_str,
            "alternatif_unvanlar": alt_unvanlar,
            "ozel_negatifler": excel_hucre_cozumle(row.get('ozel_negatifler'))
        })
    print(f"Excel'den {len(firmalar)} firma başarıyla okundu (Boş haber_unvan alanları otomatik oluşturuldu).")
except Exception as e:
    print(f"Excel okunurken hata oluştu: {e}")
    exit(1)

tum_sonuclar = []
gemini_cagrisi_sayisi, rss_toplam, arama_hatalari = 0, 0, []
indirilen_govdeler = {} # Aynı haber URL'sini birden fazla kez indirmemek için bellek içi önbellek
zaman = datetime.now(timezone.utc).isoformat()
anahtar_ifadesi = rss_anahtar_ifadesi(anahtar_kelimeler)

# 2900 Firmayı GRUP_BOYUTU (5'erli) paketler halinde tarıyoruz
toplam_grup = (len(firmalar) + GRUP_BOYUTU - 1) // GRUP_BOYUTU

for grup_idx, i in enumerate(range(0, len(firmalar), GRUP_BOYUTU), start=1):
    grup = firmalar[i:i + GRUP_BOYUTU]
    
    # Gruptaki firmaların haber_unvan'larını OR ile birleştir
    grup_unvan_sorgusu = " OR ".join([f'"{f["haber_unvan"]}"' for f in grup if f["haber_unvan"]])
    sorgu = f'({grup_unvan_sorgusu}) ({anahtar_ifadesi})'
    
    limit_degeri = MAX_HABER_FIRMA if len(grup) == 1 else MAX_HABER_GRUP
    haberler, hata = haber_ara(sorgu, limit_degeri)
    
    if hata:
        arama_hatalari.append({"firma": grup_unvan_sorgusu, "hata": hata})
        time.sleep(BEKLEME_RSS_SANIYE)
        continue
    
    rss_toplam += len(haberler)
    
    for haber in haberler:
        # Gruba gelen her haberi gruptaki firmalarla tek tek değerlendir
        for firma in grup:
            # 1. AŞAMA: Gövdeyi indirmeden ÖNCE sadece Başlık ve RSS Özeti ile hızlı ön kontrol yap
            on_haber = {
                "baslik": haber.get("baslik", ""),
                "url": haber.get("url", ""),
                "tarih": haber.get("tarih", ""),
                "rss_aciklama": haber.get("rss_aciklama", ""),
                "govde": "",
                "govde_durumu": "ATLANDI"
            }
            on_kural = kural_on_degerlendirme(firma, on_haber)
            
            # Eğer gruptan gelen bu haberin başlığında veya özetinde bu firmanın adı HİÇ geçmiyorsa atla
            if not on_kural["firma_eslesmesi_var"]:
                continue
                
            # 2. AŞAMA: Veritabanı Önbellek (Cache) Kontrolü (Gövde indirmeden önce URL bazlı kontrol)
            url_cached = cached_gemini_result(firma["firma_unvan"], haber["url"])
            
            # Sadece firma + risk eşleşmesi varsa, jenerik değilse ve daha önce DB'ye kaydedilmemişse gövdeyi indir
            aday_mi = on_kural["firma_eslesmesi_var"] and on_kural["risk_eslesmesi_var"] and not on_kural["marti_jenerik"]
            
            if url_cached and url_cached.get("govde"):
                on_haber["govde"] = url_cached["govde"]
                on_haber["govde_durumu"] = url_cached.get("govde_durumu") or "CACHE"
            elif HABER_GOVDESI_INDIR and aday_mi and not url_cached:
                if haber["url"] in indirilen_govdeler:
                    on_haber["govde"], on_haber["govde_durumu"] = indirilen_govdeler[haber["url"]]
                else:
                    govde, govde_durumu = haber_govdesi_getir(haber["url"])
                    indirilen_govdeler[haber["url"]] = (govde, govde_durumu)
                    on_haber["govde"], on_haber["govde_durumu"] = govde, govde_durumu
                    time.sleep(BEKLEME_SAYFA_SANIYE)
            
            # Gövde eklendikten sonra nihai kural değerlendirmesini ve içerik hash'ini hesapla
            kural = kural_on_degerlendirme(firma, on_haber)
            content_hash = sha256_text(on_haber.get("baslik", "") + "\n" + on_haber.get("rss_aciklama", "") + "\n" + on_haber.get("govde", ""))
            
            gemini = {"sinif": "ON_FILTREDE_ELENDI", "gerekce": "", "guven": 0.0, "model": "", "hata": ""}
            should_call = kural["firma_eslesmesi_var"] and kural["risk_eslesmesi_var"] and not kural["marti_jenerik"]
            
            # Tam hash veya URL üzerinden cache kontrolü
            cached = cached_gemini_result(firma["firma_unvan"], on_haber["url"], content_hash) or url_cached

            if should_call and cached:
                gemini.update(cached)
                gemini["cache"] = True
            elif should_call and gemini_cagrisi_sayisi < MAX_GEMINI_CAGRISI:
                gemini = gemini_degerlendir(firma, on_haber, GEMINI_MODEL_ISTENEN, GEMINI_MODEL_YEDEK, MODEL_YEDEK_KULLAN)
                gemini_cagrisi_sayisi += 1
                gemini["cache"] = False
                time.sleep(4.5)
            elif should_call:
                gemini.update({"sinif": "GEMINI_LIMITI", "gerekce": "Limit aşıldı."})

            if kural["marti_jenerik"]: durum="ELENDI_JENERIK_BAGLAM"
            elif gemini.get("cache"): durum="CACHE_YUKLENDI"
            elif gemini["sinif"]=="OLUMLU_ADAY" and gemini.get("firma_eslesme_durumu")=="ESLESIYOR" and gemini.get("olumsuz_olay"): durum="GEMINI_ADAY_INCELE"
            elif gemini["sinif"]=="ILGISIZ": durum="GEMINI_ILGISIZ"
            elif gemini["sinif"]=="BELIRSIZ": durum="GEMINI_BELIRSIZ"
            elif gemini["sinif"]=="INCELEME_GEREKLI": durum="INCELEME_GEREKLI"
            elif gemini["sinif"] in ["API_HATASI","GEMINI_LIMITI"]: durum=gemini["sinif"]
            elif not should_call: durum="ON_FILTREDE_ELENDI"
            else: durum="INCELEME_GEREKLI"

            tum_sonuclar.append({
                "firma_unvan":firma["firma_unvan"],"baslik":on_haber.get("baslik",""),"url":on_haber.get("url",""),"tarih":on_haber.get("tarih",""),"rss_aciklama":on_haber.get("rss_aciklama",""),"govde":on_haber.get("govde",""),"govde_durumu":on_haber.get("govde_durumu",""),
                "firma_eslesmeleri":", ".join(kural["firma_eslesmeleri"]),"risk_kelimeleri":", ".join(kural["risk_kelimeleri"]),"negatif_eslesmeleri":", ".join(kural["negatif_eslesmeleri"]),"kural_skoru":kural["kural_skoru"],
                "firma_eslesme_durumu":gemini.get("firma_eslesme_durumu","DOGRULANAMADI"),"haberde_gecen_firma":gemini.get("haberde_gecen_firma",""),"firma_eslesme_kaniti":gemini.get("firma_eslesme_kaniti",""),"firma_eslesme_gerekcesi":gemini.get("firma_eslesme_gerekcesi",""),"firma_eslesme_guveni":gemini.get("firma_eslesme_guveni",0),
                "olumsuz_olay":gemini.get("olumsuz_olay",False),"olay_turu":gemini.get("olay_turu","BELIRSIZ"),"olay_ozeti":gemini.get("olay_ozeti",""),"etkilenen_taraf":gemini.get("etkilenen_taraf",""),"olayin_durumu":gemini.get("olayin_durumu","BELIRSIZ"),"ciddiyet":gemini.get("ciddiyet","BELIRSIZ"),"hedef_firmaya_etki":gemini.get("hedef_firmaya_etki","BELIRLENEMEDI"),"olay_gerekcesi":gemini.get("olay_gerekcesi",""),"olay_guveni":gemini.get("olay_guveni",0),
                "gemini_sinifi":gemini["sinif"],"gemini_gerekce":gemini["gerekce"],"gemini_guven":gemini["guven"],"gemini_model":gemini.get("model",""),"nihai_durum":durum,"icerik_hash":content_hash,"analiz_surumu":ANALIZ_SURUMU
            })
            
    time.sleep(BEKLEME_RSS_SANIYE)

# Aynı çalışma içinde aynı (firma_unvan, url) çifti birden fazla kez eşleştiyse tekilleştir
tekil_sonuclar = {}
for r in tum_sonuclar:
    tekil_sonuclar[(r["firma_unvan"], r["url"])] = r
tum_sonuclar = list(tekil_sonuclar.values())

# -------- 5. VERİTABANINA YAZMA --------
aday_sayisi = sum(1 for r in tum_sonuclar if r["nihai_durum"] == "GEMINI_ADAY_INCELE")

cursor.execute("""
INSERT INTO arama_gecmisi (sorgu_zamani, firma_sayisi, rss_sonucu, aday_sayisi, gemini_cagrisi, hata_sayisi)
VALUES (%s, %s, %s, %s, %s, %s) RETURNING arama_id
""", (zaman, len(firmalar), rss_toplam, aday_sayisi, gemini_cagrisi_sayisi, len(arama_hatalari)))
arama_id = cursor.fetchone()[0]

for r in tum_sonuclar:
    cursor.execute("""
    INSERT INTO haber_eslesmeleri (firma_unvan,url,baslik,haber_tarihi,rss_aciklama,govde,govde_durumu,firma_eslesmeleri,risk_kelimeleri,negatif_eslesmeleri,kural_skoru,gemini_sinifi,gemini_gerekce,gemini_guven,nihai_durum,gemini_model,icerik_hash,ilk_gorulme,son_gorulme,son_arama_id,firma_eslesme_durumu,haberde_gecen_firma,firma_eslesme_kaniti,firma_eslesme_gerekcesi,firma_eslesme_guveni,olay_turu,olay_ozeti,etkilenen_taraf,olayin_durumu,ciddiyet,hedef_firmaya_etki,olay_gerekcesi,olay_guveni,analiz_surumu)
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    ON CONFLICT (firma_unvan,url) DO UPDATE SET
      baslik=EXCLUDED.baslik,haber_tarihi=EXCLUDED.haber_tarihi,rss_aciklama=EXCLUDED.rss_aciklama,govde=CASE WHEN EXCLUDED.govde IS NOT NULL AND EXCLUDED.govde!='' THEN EXCLUDED.govde ELSE haber_eslesmeleri.govde END,govde_durumu=EXCLUDED.govde_durumu,firma_eslesmeleri=EXCLUDED.firma_eslesmeleri,risk_kelimeleri=EXCLUDED.risk_kelimeleri,negatif_eslesmeleri=EXCLUDED.negatif_eslesmeleri,kural_skoru=EXCLUDED.kural_skoru,gemini_sinifi=EXCLUDED.gemini_sinifi,gemini_gerekce=EXCLUDED.gemini_gerekce,gemini_guven=EXCLUDED.gemini_guven,nihai_durum=EXCLUDED.nihai_durum,gemini_model=EXCLUDED.gemini_model,icerik_hash=EXCLUDED.icerik_hash,son_gorulme=EXCLUDED.son_gorulme,son_arama_id=EXCLUDED.son_arama_id,firma_eslesme_durumu=EXCLUDED.firma_eslesme_durumu,haberde_gecen_firma=EXCLUDED.haberde_gecen_firma,firma_eslesme_kaniti=EXCLUDED.firma_eslesme_kaniti,firma_eslesme_gerekcesi=EXCLUDED.firma_eslesme_gerekcesi,firma_eslesme_guveni=EXCLUDED.firma_eslesme_guveni,olay_turu=EXCLUDED.olay_turu,olay_ozeti=EXCLUDED.olay_ozeti,etkilenen_taraf=EXCLUDED.etkilenen_taraf,olayin_durumu=EXCLUDED.olayin_durumu,ciddiyet=EXCLUDED.ciddiyet,hedef_firmaya_etki=EXCLUDED.hedef_firmaya_etki,olay_gerekcesi=EXCLUDED.olay_gerekcesi,olay_guveni=EXCLUDED.olay_guveni,analiz_surumu=EXCLUDED.analiz_surumu
    """,(r["firma_unvan"],r["url"],r["baslik"],r["tarih"],r["rss_aciklama"],r["govde"],r["govde_durumu"],r["firma_eslesmeleri"],r["risk_kelimeleri"],r["negatif_eslesmeleri"],int(r["kural_skoru"]),r["gemini_sinifi"],r["gemini_gerekce"],float(r["gemini_guven"]),r["nihai_durum"],r["gemini_model"],r["icerik_hash"],zaman,zaman,arama_id,r["firma_eslesme_durumu"],r["haberde_gecen_firma"],r["firma_eslesme_kaniti"],r["firma_eslesme_gerekcesi"],float(r["firma_eslesme_guveni"]),r["olay_turu"],r["olay_ozeti"],r["etkilenen_taraf"],r["olayin_durumu"],r["ciddiyet"],r["hedef_firmaya_etki"],r["olay_gerekcesi"],float(r["olay_guveni"]),r["analiz_surumu"]))
conn.commit()
cursor.close()
conn.close()

# -------- 6. E-POSTA BİLDİRİMİ --------
aday_haberler = [r for r in tum_sonuclar if r["nihai_durum"] == "GEMINI_ADAY_INCELE"]

msg = MIMEMultipart("alternative")
msg["From"] = GONDEREN_MAIL
#msg["To"] = ", ".join(ALICI_MAILLER)
msg["To"] = "undisclosed-recipients:;"

if aday_haberler:
    df_rapor = pd.DataFrame(aday_haberler)[["firma_unvan","haberde_gecen_firma","firma_eslesme_durumu","firma_eslesme_guveni","olay_turu","olay_ozeti","olayin_durumu","ciddiyet","hedef_firmaya_etki","gemini_gerekce","url"]].copy()
    df_rapor.columns=["Hedef Firma","Haberde Geçen Firma","Firma Eşleşmesi","Eşleşme Güveni","Olay Türü","Olay Özeti","Olay Durumu","Ciddiyet","Hedef Firmaya Etki","LLM Yorumu","Haber"]
    df_rapor["Eşleşme Güveni"]=df_rapor["Eşleşme Güveni"].map(lambda x:f"{float(x or 0):.0%}")
    df_rapor["Haber"]=df_rapor["Haber"].apply(lambda x:f'<a href="{html.escape(str(x),quote=True)}">Habere Git</a>')
    html_tablo=df_rapor.to_html(escape=False,index=False)
    
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
