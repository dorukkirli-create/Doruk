# Aylik yansitma isleyisi (kisisel veri icermeyen ozet)

Bu dosya, RHI Rusya icin iki aylik yansitma akisinin kural setini tasir. Isimli
ornekler, karar defterleri ve 1C'den turetilmis sozlukler repo DISINDA, ayri bir
baglam paketinde tutulur (repo herkese aciktir).

## Iki ayri akis

| Akis | Girdi | Marj | Cikti |
|---|---|---|---|
| Seyahat acentesi (bilet, otel, vize, bagaj) | acentenin aylik listesi + cari dokumu | %3 | calisma Excel + "Resmi faturalar icin rusca data" |
| Ankara ofisi yansitmasi (egitim, assessment, hukuk, sigorta, ISG, danismanlik) | yansitma maili + fatura ekleri | %20 | dagilim Excel + proje basina Schet / Akt / Otchet + Invoices |

## Seyahat akisi kurallari
1. Iade satirlari negatife cevrilir; kolon toplami acentenin ODENECEK rakamina esit olmali.
2. Kapsam disi satirlar (transfer, kargo, celenk, toplanti organizasyonu, rezervasyon ucreti) resmiye girmez; kisisiz olanlar RHI NAKIT, ucuncu kisi odemeleri EFT.
3. Her satira sicil ve Kiril isim: 1C ay sonu listesi (Kiril kolonu) esas; aile/refakatci yolculara calisanin sicili, kendi Kiril adi. 1C'de olmayanlar 1C token sozlugunden cevrilir ve isaretlenir.
4. %3 yalniz resmiye giren satirlarda; EFT/NAKIT/karar satirlarinda %3 kolonu bos.
5. Podrazdelenie onceligi: listedeki santiye kolonu > onceki ayin ayni kisiye verdigi deger > 1C projesi. Bagajda santiye/talep eden birim esas, bolme notu uygulanmaz.
6. Resmi dosya sozlesme sirasiyla gruplanir; ayni bilet numarasi (yoksa ayni kisi + guzergah, ayni ay, onceki ayda ayni bilet yoksa) iadeleri netlenir; onceki ayin iadesi resmiye girmez, nakit bakiyeden dusulur.
7. Bilet numaralari acentenin cari dokumundeki "Evrak No" alanindan gelir (tarih + tutar + isim ile birlestirme).
8. Yerel ayara bagli TEXT(...,"ДД.ММ.ГГГГ") formulu kullanilmaz; tarih metni duz yazilir.

## Ankara ofisi akisi kurallari
1. Turk faturalari kendi tarihinin kuruyla USD'ye cevrilir, sonra x1,20 (ISG doktor ve isveren maliyeti kalemleri haric).
2. Her gider ailesi icin kisi bazli sayfa; OZET podrazdelenie x aile pivotu; ODEME grubu (4 proje) SUMIF ile.
3. Resmi belgeler proje grubu basina uc dosya: Schet (kalemler), Akt (ayni kalemler), Otchet (kisi basina saat x birim fiyat).
   100 $/saat: assessment, hukuk, egitim. 120 $/saat: ISG (0,5 s), proje yonetimi danismanligi. Ferdi kaza sigortasi tek satir, adet 1, fiyat = plug.
4. Resmi fatura toplami gercek maliyetin onluga yuvarlanmis halidir (kurus kalmasin diye); sigorta satiri farki emer.
5. Kalem eslemesi: IK yonetimi = assessment + egitimler; hukuk = arabuluculuk; saglik = sigorta + ISG; proje yonetimi = danismanlik isveren maliyeti.

## Teslim biçimi
- Seyahat: iki loose Excel (+ acentenin carisi dokunulmadan).
- Ankara ofisi: dagilim xlsx + "Первичные документы.7z" + "Invoices.7z".
- Zip'ler UTF-8 bayrakli yazilir (Python zipfile); kabuk `zip` Kiril dosya adlarini Windows'ta bozar.

## Finans ekibinin final duzeltmeleri (seyahat akisi, Agustos 2026)
Finans ekibi hazirladigimiz iki Excel'i yapiyi degistirmeden acenteye gonderdi. Yaptigi kucuk duzeltmeler kural olarak koda alindi:
1. %3 kolonu deger degil formul yazilir (`=+F2*1.03`).
2. Ozet blokta o ay satiri olmayan firmalar ve bos EFT satiri yazilmaz; blok yalniz dolu firmalar + RHI NAKIT + TOPLAM.
3. Ayri sozlesmesi olmayan istirak (BSA) calisanlari dogrudan merkez ofis podrazdeleniesine yazilir; ayri proje etiketi acilmaz.
4. Gydan isleri "ALNG2-GBS" proje etiketiyle gider; sozlesme yine merkez sozlesmesi.
5. Podrazdelenie kararinda finans ekibinin kisiye ozel gecmis kararlari 1C projesinin onune gecer (sicil bazli istisna listesi).
6. Kiril yazimda Turkce s/i/o harfleri ш/ы/ё ile karsilanir; ucuncu ad (baba adi disinda) yazilmaz.
7. Resmi dosyada otel adindan semt/mahalle eki atilir.
Uretim kodu bu kurallarla finans ekibinin gonderdigi dosyayi hucre hucre yeniden uretir.

