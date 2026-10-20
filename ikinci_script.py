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
