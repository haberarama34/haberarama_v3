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
