# Masraf Merkezi Otomasyonu

Tedarikçi faturalarındaki her satırı, o satırdaki **kişiye** göre doğru
**masraf merkezine** (proje / şantiye) mahsuplaştıran masaüstü aracı.

## Ne işe yarar

Seyahat acentesi, eğitim ve sağlık faturaları RHI'ya tek bir toplam olarak
geliyor; içindeki satırların hangi projeye yazılacağı fatura üzerinde yazmıyor,
sadece kişi adı geçiyor. Finans ekibi bugün bu dağıtımı her ay elle yapıyor:
130 satırlık bir seyahat faturası için 130 kez "bu kişi hangi şantiyede
çalışıyordu" sorusunu cevaplamak gerekiyor.

Bu araç aynı işi personel ana verisiyle otomatik yapar. Emin olduğu satırları
doğrudan dağıtır, emin olmadıklarını gerekçesiyle birlikte inceleme kuyruğuna
koyar. Amaç elle çalışmayı sıfırlamak değil, **elle bakılacak satır sayısını
azaltmak** ve her kararın nedenini görünür kılmak.

Ölçülen durum (Temmuz 2026 seyahat faturası, 134 satır): 82 satır otomatik
dağıtılıyor, 52 satır insana kalıyor. Bu 52'nin 20'si zaten grup şirketi
personelidir ve **hiçbir zaman** otomatik dağıtılamaz — onlar için doğru
davranış "eşleşmedi" demektir. Sayılar ve nasıl üretildikleri aşağıda
[Ölçülen performans](#ölçülen-performans) bölümündedir.

## Nasıl çalışır

İki temel fikir var.

**1. Kademeli eşleştirme.** Fatura metninden çıkarılan isim, en güvenilirden
en zayıfa doğru sıralanmış kademelerden geçirilir. İlk başarılı olan kademe
kazanır ve o kademenin güven skoru sonuca yazılır. Hiçbiri tutmazsa satır
"eşleşmedi" olarak işaretlenir; sistem **tahmin yürütmez**.

**2. Tarihe göre dönem seçimi.** Personel ana verisi aylık snapshot'lardan
oluşur; aynı kişinin her ay için bir satırı vardır. Kişilerin %1,4'ünün görev
yeri dönemler arasında değişir. Bu yüzden masraf merkezi, faturanın kesildiği
**aya ait** kayıttan okunur, en son kayıttan değil. Şubat'ta GPP'de olup
Temmuz'da Amur'a geçen biri için Şubat faturası GPP'ye yazılır.

Akış:

```
  Fatura dosyası (.xls/.xlsx/.csv/.msg)
            |
            v
  [1] Dosya tipini tanı  -----> antik_cari | yuzyil_dagitilmis | energo_* | koc | genel
            |
            v
  [2] Satırları oku, açıklamadan kişi adını ayıkla
            |               (otel adı, güzergah kodu, bilet no, tarih temizlenir)
            v
  [3] Kademeli eşleştirme
            |
            +--> sicil no verilmiş mi?            --> EŞLEŞTİ  (1,00)
            +--> TC kimlik köprüsü tutuyor mu?    --> EŞLEŞTİ  (0,99)
            +--> aliases.csv'de var mı?           --> EŞLEŞTİ  (0,98)
            +--> tam isim tek kişide mi?          --> EŞLEŞTİ  (0,95)
            +--> isim alt kümesi tek kişide mi?   --> EŞLEŞTİ  (0,90)
            +--> bitişik ad açılıyor mu?          --> EŞLEŞTİ  (0,92)
            |        (MUSTAFAKEMAL -> MUSTAFA KEMAL)
            +--> transliterasyon varyantı tutuyor mu? --> İNCELE  (0,88)
            |        (IYLMAZ GEKHAN -> YILMAZ GOKHAN)
            +--> kesik isim öneki tutuyor mu?     --> İNCELE   (0,85)
            +--> ek kişi defterinde mi?           --> İNCELE   (0,70)
            +--> bulanık benzerlik yeterli mi?    --> İNCELE   (0,79-0,90)
            +--> soyadı bir çalışanla aynı mı?    --> İNCELE   (0,60) "aile bireyi"
            +--> hiçbiri                          --> EŞLEŞMEDİ (0,00)
            |
            v
  [4] Fatura tarihine en yakın DÖNEM kaydını seç
            |
            v
  [5] Görev yeri -> masraf merkezi kodu (masraf_merkezi_haritasi.csv)
            |
            v
  [6] Excel çıktısı: Sonuç | İncele | Eşleşmedi | Özet
```

Yukarıdaki "EŞLEŞTİ" yalnızca **0,90 ve üstü** kademeler için otomatik kabul
demektir; altındakiler kişiyi bulur ama kararı insana bırakır.

Birden fazla aday çıkarsa güven skoru 0,58'in üstüne **çıkamaz** ve sicil
doldurulmaz; satır adaylarıyla birlikte incelemeye düşer. Tek istisna: bütün
adaylar aynı görev yerindeyse masraf merkezi zaten tektir, öneri verilir ama
yine de doğrulanması istenir.

Çalışma zamanında yapay zeka veya internet kullanılmaz. Her şey deterministik
Python'dur; aynı girdi her zaman aynı çıktıyı verir ve sonuç ofis dışında,
çevrimdışı bir dizüstünde de aynıdır.

## Kurulum

**Windows:** `baslat.bat` dosyasına çift tıklayın. İlk çalıştırmada sanal ortam
kurulur ve paketler indirilir (birkaç dakika), sonraki açılışlar hızlıdır.

**Linux / macOS:**

```bash
chmod +x baslat.sh    # sadece ilk seferde
./baslat.sh
```

Her ikisi de tarayıcıda `http://localhost:8501` adresini açar. Kapatmak için
konsol penceresinde `Ctrl+C`.

Elle kurmak isterseniz:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Gereken: Python 3.11, `pandas`, `openpyxl`, `xlrd`, `rapidfuzz`, `streamlit`,
`xlsxwriter`, `extract-msg`.

## Kullanım adımları

1. **Personel verisini yükleyin.** *Ayarlar* sekmesinde
   `2025_2026_giris_cikis.xlsx` dosyasının yolunu verin. Dosya 24 MB'dır ve ilk
   okunuşu yaklaşık 25 saniye sürer; sonrasında yanına bir önbellek dosyası
   yazılır ve açılışlar 2-3 saniyeye iner. Dosya değişince önbellek kendini
   yeniler.
2. **Fatura dosyalarını bırakın.** *Fatura İşle* sekmesinde dosyaları
   sürükleyip bırakın veya bir klasör yolu verin. Birden fazla dosya aynı anda
   işlenebilir; Outlook `.msg` dosyalarının ekleri otomatik açılır.
3. **İşleyin.** Dosya tipi otomatik tanınır, doğru okuyucu seçilir. 130 satırlık
   bir dosya birkaç saniye sürer.
4. **İnceleyin.** *İnceleme* sekmesinde düşük güvenli satırlar tek tek gelir.
   Her satırın yanında sistemin **neden** o sonuca vardığı Türkçe yazar ve
   adaylar listelenir. Doğru kişiyi seçtiğinizde sistem bunu öğrenir.
5. **Excel'i indirin.** Dört sayfalı çıktı: `Sonuç` (tüm satırlar),
   `İncele` (elle bakılacaklar), `Eşleşmedi` (kişi bulunamayanlar), `Özet`
   (dağılımlar, oranlar, tutar toplamları).

## Desteklenen dosya tipleri

| Kaynak | Tanıma ipucu | Kimlik anahtarı | Masraf merkezi |
|---|---|---|---|
| Antik / Yüzyıl ham cari hareket dökümü (`.xls`) | "Cari Hareket Dökümü Detay" başlığı, `İşlem`/`Evrak No`/`Borç` kolonları | Kişi adı açıklama metnine gömülü | Yok, eşleştirmeden gelir |
| Yüzyıl elle dağıtılmış (`.xlsx`) | `S.NO`, `AÇIKLAMA`, `ŞANTİYESİ` kolonları | Kişi adı açıklamada | Var (`ŞANTİYESİ`) — doğruluk referansı |
| Energo assessment yansıtma (`.xlsx`) | `Fatura Detay` + `Kişi Listesi` sayfaları, `Katılımcı` kolonu | Ad Soyad | Yok |
| Energo arabuluculuk (`.xlsx`) | `PERSONEL T.C.`, `PROJE` kolonları | **TC kimlik no** | Var (`PROJE`) |
| Sağlık kontrol listesi (`.xlsx`) | `BORDROLU LİSTE` sayfası, `TCKN` + `ŞANTİYE` | **TC kimlik no** + doğum tarihi | Var (`ŞANTİYE`) |
| Koç Üniversitesi katılımcı listesi (`.xlsx`) | `ID`, `Ad Soyad`, `Katılım Tarihi` | **Sicil numarası** (`ID` kolonu) | Yok |
| Outlook e-postası (`.msg`) | Dosya uzantısı | Ekteki dosyaya göre | Ekteki dosyaya göre |
| Tanınmayan tablo (`.xlsx`, `.csv`) | Genel okuyucu, kolon adlarından çıkarım | Bulunabilene göre | Varsa okunur |

En kolay eşleşen kaynaklar sicil veya TC kimlik taşıyanlardır. Seyahat
faturaları en zorudur: kimlik alanı yoktur, sadece serbest metinde isim vardır.

## Güven skorları ne demek

| Yöntem | Güven | Ne zaman oluşur |
|---|---|---|
| `sicil` | 1,00 | Kaynak dosyada sicil numarası var (Koç katılımcı listesi gibi) |
| `tckn` | 0,99 | TC kimlik `veri/tckn_sicil.csv` köprüsünde tek bir sicile bağlanıyor |
| `alias` | 0,98 | Bu ismi daha önce siz elle onaylamışsınız (`veri/aliases.csv`) |
| `tam_isim` | 0,95 | Normalize isim personel verisinde **tek** kişiye denk geliyor |
| `tam_isim` (bitişik ad) | 0,92 | `MUSTAFAKEMAL` sözlükle `MUSTAFA KEMAL` olarak açıldı, sonuç tek kişi |
| `alt_kume` | 0,90 | Fatura ismi personel isminin alt kümesi, tek aday (ikinci ad eksik) |
| `transliterasyon` | 0,88 | Rusça transliterasyon geri çevrildi (`GEKHAN` -> `GOKHAN`), tek aday |
| `prefix` | 0,85 | İsim bilet sisteminde kesilmiş, önek tek kişiye uyuyor |
| `alt_kume` (zayıf) | 0,72 | Alt küme eşleşmesi ama isim çok yaygın |
| `ek_defter` | 0,70 | Kişi ana veride yok, yardımcı listelerden (sağlık listesi vb.) bulundu |
| `bulanik` | 0,79-0,90 | Yazım hatası toleranslı benzerlik (rapidfuzz ≥ 88 puan) ve ikinci adayla arada en az 6 puan fark var |
| `aile` | 0,30-0,60 | Soyadı aynı dosyada kesin eşleşen bir çalışanla aynı; eş/çocuk olabilir. Kanıtın gücüne göre dört kademe (aynı dosyada soyadaş + tek görev yeri en güçlüsü) |
| çoklu aday | ≤ 0,58 | Aynı isimde birden fazla çalışan var; sicil **doldurulmaz** |
| `yok` | 0,00 | Hiçbir kademe tutmadı |

Karar eşikleri (`masraf/masraf_merkezi.py`): **0,90 ve üstü** güvene sahip
*ve* hiç uyarı taşımayan satırlar `Sonuç` sayfasında otomatik kabul edilir
(`GUVEN_ESIGI = 0.90`). **0,50'nin altındakiler** eşleşmemiş sayılır ve
`Eşleşmedi` sayfasına gider (`ALT_ESIK = 0.50`). Arada kalan her şey `İncele`
sayfasına düşer. Eşikler *Ayarlar* sekmesinden değiştirilebilir.

Buradan çıkan sonuç önemlidir: `transliterasyon` (0,88), `prefix` (0,85),
`bulanik` (0,79-0,90), `ek_defter` (0,70) ve `aile` (0,60) kademelerinin
**hiçbiri kendi başına otomatik kabul edilmez** — hepsi insan onayına gider.
Otomatik kabul yalnızca sicil, TC kimlik, onaylanmış alias, tek kişiye denk
gelen tam isim ve bitişik ad açılımı kademelerinden gelir.

Bir satır yüksek güvenle eşleşse bile uyarı taşıyorsa (kişi belge tarihinden
önce işten ayrılmış, görev yeri haritada yok, tüzel kişi çelişkisi) otomatik
kabul edilmez ve incelemeye düşer.

## Öğrenme

Sistem kullandıkça iyileşir ama içinde yapay zeka **yoktur**. Öğrenme
tamamen tablo doldurmaktır:

- İnceleme ekranında bir satırı bir sicile bağladığınızda, normalize isim ve
  sicil `veri/aliases.csv` dosyasına yazılır. Aynı isim bir daha geldiğinde
  0,98 güvenle, hiçbir tahmin yapılmadan eşleşir.
- "Bu kişi çalışanımız değil" dediğinizde kayıt `veri/harici_kisiler.csv`
  dosyasına gider ve bir daha inceleme kuyruğunu şişirmez.
- Sağlık kontrol listesi, katılımcı listesi gibi yardımcı dosyalar
  işlendiğinde içlerindeki kişiler `veri/ek_kisiler.csv` defterine, TC kimlik
  bilgileri `veri/tckn_sicil.csv` köprüsüne otomatik eklenir.

Bu yüzden model eğitimi, internet bağlantısı veya API anahtarı gerekmez.
Bir eşleşme neden kurulduğu sorulduğunda cevap her zaman bir CSV satırıdır;
"model böyle karar verdi" gibi bir cevap yoktur. Dosyaları bir metin
düzenleyiciyle açıp elle de düzeltebilirsiniz.

## Ölçülen performans

Aşağıdaki sayıların hepsi **ölçülmüştür**, tahmin değildir. Kendiniz
üretebilirsiniz:

```bash
python3 testler/kapsam_olc.py      # kapsam ve otomasyon oranı
python3 -m testler.dogruluk_olc    # elle dağıtılmış dosyaya karşı doğruluk
```

Ölçüm örneklemi: Temmuz 2026 seyahat faturası, Mayıs–Temmuz 2026 Energo
yansıtma dosyaları ve Koç Üniversitesi katılımcı listesi. Her iki betik de
öğrenen defterlerin geçici bir kopyasıyla çalışır; ölçüm `veri/` dizinini
kirletmez ve her koşuda aynı noktadan başlar (iki ardışık koşu birebir aynı
sonucu verdi).

### Dosya bazında otomasyon oranı

| Dosya | Tip | Satır | Kişi çıkarıldı | Sicil bulundu | OTOMATİK | İNCELE | EŞLEŞMEDİ | Otomasyon |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| `ANTIK_CARI_TEMMUZ_2026.xls` | antik_cari | 134 | 132 | 105 | 82 | 32 | 20 | **%61,2** |
| `YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx` | yuzyil_dagitilmis | 134 | 132 | 105 | 79 | 35 | 20 | **%59,0** |
| `ASSESSMENT_YANSITMA_2026_05_06.xlsx` | energo_assessment | 6 | 6 | 5 | 4 | 1 | 1 | **%66,7** |
| `ARABULUCULUK_2026_06_07.xlsx` | energo_arabulucu | 25 | 25 | 22 | 0 | 25 | 0 | **%0,0** |
| `SAGLIK_KONTROL_LISTE.xlsx` | energo_saglik | 50 | 50 | 34 | 14 | 34 | 2 | **%28,0** |
| `KOC_UNI_KATILIMCI_LISTESI.xlsx` | koc_katilimci | 50 | 50 | 50 | 46 | 4 | 0 | **%92,0** |
| **TOPLAM** | | **399** | **395** | **321** | **225** | **131** | **43** | **%56,4** |

Çözülen (OTOMATİK + İNCELE) satır oranı **%89,2**. Kişi çıkarılamayan 4 satırın
tamamı gerçekten kişisiz kurumsal giderdir (cenaze çelengi, toplantı
organizasyonu); bunlar kişiye mahsuplaşmaz.

Uçlardaki iki sayı tesadüf değil, doğrudan **kaynak dosyada kimlik alanı olup
olmadığını** ölçüyor:

- **Koç katılımcı listesi %92** — dosyada `ID` kolonu doğrudan sicil numarası.
- **Arabuluculuk %0** — dosyada TC kimlik var ama `veri/tckn_sicil.csv` köprüsü
  boş olduğu için 25 satırın hepsi isim üzerinden bulunup 0,70 güvenle
  incelemeye düşüyor. Köprü doldurulunca bu dosya 0,99 güvenle otomatiğe geçer.
  Aynı sebep sağlık listesinin %28'ini de açıklar. **Bu iki dosyadaki düşük
  oran algoritma zayıflığı değil, eksik bir yapılandırma dosyasıdır.**

`ornek_mail.msg` bilerek toplama **katılmaz**: ekleri yukarıdaki faturaların
aynısını ikinci kez taşır, toplama girseydi her satır iki kez sayılır ve oran
şişerdi (ilk ölçümde toplam 804 görünüyordu, gerçek 399). Ekler ayrı bir
tabloda raporlanır ve satır satır aynı sonucu verir — yani `.msg` okuyucusu
doğru çalışıyor.

### Eşleşmeyen 43 satırın nedeni

| Kategori | Adet | Pay | Anlamı |
|---|---:|---:|---|
| `ALGORITMA` | **0** | %0,0 | Kişi veride var, eşleştirici bulamadı — **düzeltilebilir** |
| `PARSER` | **0** | %0,0 | Satırda kişi var ama çıkarılamadı — **düzeltilebilir** |
| `VERI_KAPSAMI` | 39 | %90,7 | Kişi personel ana verisinde yok — düzeltilemez |
| `KISISIZ` | 4 | %9,3 | Satırda kişi yok (kurumsal gider) — kusur değil |

Düzeltilebilir kusur **sıfır**. Eşleşmeyen her satır ya grup şirketi / dış
danışman / taşeron personelidir (veri kapsamı dışı), ya da kişiye
mahsuplaşmayan bir kurumsal giderdir.

### Doğruluk: elle dağıtılmış dosyaya karşı

Temmuz 2026'nın elle dağıtılmış hali (`YUZYIL_..._ELLE_DAGITILMIS.xlsx`) ile
otomasyonun çıktısı 134 satırın **134'ünde** tutar + tarih ile hizalandı;
hizalama hatası yok.

**Kritik bulgu: elle dosya "doğruluk referansı" değildir.** Naif karşılaştırma
%48,6 verdi. Bu 57 "hatanın" nedeni tek tek açıldığında hatanın otomasyonda
olmadığı görüldü — iki dosya **aynı soruyu cevaplamıyor**:

```
GPP Project çalışanı -> elle 'RHI'            : 54 satır
GPP Project çalışanı -> elle 'UST LUGA GPP'   : 12 satır
```

Aynı proje, aynı ay, iki farklı etiket. Kişi bazında çelişki **yok** (aynı kişi
her zaman aynı etiketi almış), yani insan tutarsız değil. Ayrımı yapan değişken
kişi değil, **biletin güzergahı**:

| Güzergah sınıfı | Proje/şirket yazılmış | 'RHI' (varsayılan) |
|---|---:|---:|
| ESB kalkışlı giriş | **27** | 2 |
| Güzergahsız | 2 | 28 |
| Diğer giriş | 1 | 13 |
| Çıkış | 1 | 39 |
| Gidiş-dönüş | 0 | 18 |

Ankara (ESB) kalkışlı tek yön biletler toplu mobilizasyondur ve masrafı
üstlenen projeye yazılmış; diğer her şey merkeze bırakılmış. Yani `ŞANTİYESİ`
kolonu **"bu kişi hangi projede çalışıyor"** sorusunu değil **"bu bileti hangi
tüzel kişi ödeyecek"** sorusunu cevaplıyor. Otomasyonun ürettiği şey
birincisidir.

Bu yüzden doğruluk üç ayrı okumayla raporlanır:

| Okuma | Sonuç | Ne ölçer |
|---|---:|---|
| `proje` (naif) | %48,6 (54/111) | Taksonomi farkını ölçer, doğruluğu değil. **Yanıltıcı.** |
| `tüzel` | %96,4 (107/111) | Otomasyon elle dosyayla çelişiyor mu. Zayıf test: 'RHI' 100 satırda ayrım yapmıyor. |
| `bilgi` | **%86,7 (13/15)** | Elle dosyanın gerçek proje bilgisi taşıdığı satırlar. **Tek geçerli ölçüm budur.** |

### Bulunan gerçek otomasyon hataları: 2 — ikisi de İNCELE bayrağıyla yakalandı

- **#131 TURAN MEHMET** (ESB-LED, 31.07). Sicil 549718: Amursky'de çalışmış ama
  **06.04.2026'da çıkmış**. Aynı günün aynı partisindeki iki kişi için hem elle
  hem otomasyon "GPP" diyor. Bu neredeyse kesinlikle yeni işe girmiş **başka bir
  Mehmet Turan**. Elle dosya doğru, otomasyon yanlış — ama sistem iki uyarıyla
  ("gider ayında personel kaydı yok", "belge tarihinden önce ayrılmış") tam
  olarak doğru yeri işaret etti ve satırı otomatik kabul etmedi.
- **#69 SERKAN KOCAK** — `aile` kuralı soyada değil **ada** takılıp Kocak Cem'e
  0,45 güvenle bağlandı. Eşik altında kaldığı için `Eşleşmedi` sayfasına düştü,
  yanlış mahsuplaşma üretmedi.

Elle dosyada **düzeltilmesi gereken bir insan hatası bulunmadı.** Uyuşmazlıkların
tamamı ya taksonomi farkı, ya veri kapsamı dışı kişi, ya da yukarıdaki iki
otomasyon hatasıdır.

## İki personel dosyası kullanın

Ana personel dosyası (`2025_2026_giris_cikis.xlsx`) aylık snapshot serisidir ve
**sadece RHI ile UST LUGA tüzel kişilerini** kapsar. Giderin yapıldığı ayın
kaydını buradan okuruz.

1C personel listesi (`1C Personnel List ...xlsx`) tek tarihlidir ama **bütün grup
şirketlerini** kapsar. Renservis, Renstroydetal, RC, One Tower, Top Tower ve BSK
personeli ancak burada bulunur.

Ölçülen katkı: 1C listesindeki 17.517 isimli kaydın 5.234'ü ana veride yoktur.
İki dosya aynı sicil uzayını kullanır (12.283 ortak sicil), bu yüzden güvenle
birlikte kullanılırlar.

| Ölçüm | Sadece ana veri | Ana veri + 1C listesi |
|---|---|---|
| Masraf merkezi çözülen satır | 362 / 405 | 397 / 405 |
| Oran | yüzde 89,4 | yüzde 98,0 |
| Hiç eşleşmeyen | 52 | 14 |

Ayarlar sekmesinde ikinci dosya yolunu da verin. Zorunlu değildir, olmadan da
çalışır; ama olmadan grup şirketi personeli bulunamaz.

1C listesinden gelen bir kayıt her zaman şu uyarıyı taşır: *"1C listesi tek
tarihli olduğu için gider ayındaki durum doğrulanamadı."* Bu kasıtlıdır. O kişi
için ay bazlı kontrol yapılamaz, karar insana bırakılır.

---

## Yeni bir format geldiğinde

Üç durum var ve üçünde de yapmanız gereken farklı.

**1. Bilinen şablon.** Altı dosya ailesi otomatik tanınır. Hiçbir şey yapmayın.

**2. Yeni tedarikçi, tanıdık kolon adları.** Genel okuyucu kolon adlarını
anahtar kelimeyle bulur. Türkçe (`Personel Adı Soyadı`, `Net Tutar`, `Proje`) ve
İngilizce (`Employee Name`, `Amount`, `Cost Center`) adların çoğu zaten tanınır.
Yine bir şey yapmanız gerekmez.

**3. Kolon adları hiç tanıdık değil.** Örneğin `Ref, Dt, Beneficiary, Note, Val`.
Uygulama sessizce boş dönmez; dosyadaki kolon adlarını listeler ve hangi alanın
eksik olduğunu söyler. Siz `veri/kolon_esanlamlilari.csv` dosyasına satır
eklersiniz:

```
alan;kolon_adi;not
kisi;Beneficiary;Yeni tedarikci X boyle yaziyor
tarih;Dt;
tutar;Val;
```

Geçerli `alan` değerleri: `kisi`, `sicil`, `tckn`, `tutar`, `tarih`, `santiye`.
Bir kez eklemek yeterlidir, sonraki bütün dosyalarda çalışır. Doğrulandı: hiç
tanınmayan kolon adlarıyla sıfır satır dönen bir dosya, üç satır eklendikten
sonra tam okundu.

**4. Dosyanın şekli tamamen farklı.** Kişi adı serbest metnin içine gömülüyse
(seyahat dökümündeki `TK4093099626 OZAKAY/MUSTAFAKEMAL MR IST-CDG BILET BEDELI`
gibi) kolon sözlüğü yetmez, o kalıp için kod yazmak gerekir. Yılda bir iki kez
karşılaşılacak bir durumdur.

---

## Bilinen kısıtlar

Bunları bilerek kullanın; araç bunları gizlemez, çıktıda uyarı olarak gösterir.

- **Personel ana verisinde TC kimlik ve pasaport numarası yoktur.** Sadece
  sicil, ad soyad ve doğum tarihi kimlik alanı olarak bulunur. TC kimlik ile
  eşleşme yapabilmek için `veri/tckn_sicil.csv` köprüsü elle veya İK
  sisteminden doldurulmalıdır. Köprü boşken arabuluculuk ve sağlık
  dosyalarındaki TCKN'ler isim üzerinden eşleşmeye çalışır.
- **Dosya yalnızca RHI ve UST LUGA tüzel kişilerini kapsar.** Renservis,
  Renstroydetal, One Tower, Top Tower, RC Peter, RC Moskova personeli bu
  veride **yoktur** ve bulunamaz. Bu kişiler eşleşmedi olarak gelir; doğru
  davranış budur, çünkü masrafları başka bir şirkete yansıtılacaktır.
- **Dönem aralığı 2025-11 ile 2026-07 arasıdır.** Bu aralığın dışında tarihli
  faturalarda en yakın dönem kullanılır ve satır uyarı taşır. Yeni aylar
  eklendikçe aralık kendiliğinden genişler.
- **Henüz işe başlamamış aday ve yeni girenler ana veride olmaz.** Bunlar için
  sağlık kontrol listesi gibi yardımcı kaynaklardan ek kişi defteri
  beslenmelidir. Ölçüm: eşleşmeyen 43 satırın 39'u (%90,7) tam olarak bu
  gruptur ve hiçbiri algoritma kusuru değildir.
- **Aile bireyleri ve dış danışmanlar otomatik eşleşmez.** Eşin veya çocuğun
  bileti çalışanın soyadıyla gelir; sistem bunu "aile bireyi olabilir" diye
  işaretler, masraf merkezini o çalışandan devralır ama **incelemeye
  gönderir**. Soyadı 8'den fazla çalışanda geçiyorsa aile varsayımı hiç
  kurulmaz, çünkü aynı soyadın tesadüf olma ihtimali yüksektir.
- **Aynı isimli çalışanlar otomatik seçilmez.** İsim çakışması genelde %3,6,
  Hint uyruklu personelde %10,9'dur (Kumar Manoj gibi çok yaygın isimler).
  Çakışan satır adaylarıyla birlikte incelemeye düşer.
- **Bordrosuz taşeron kayıtları isim eşleştirmesine girmez.** Ana veride
  44.482 satırın adı boştur (sahte sicil numaralarıyla). Bunlar sicil
  indeksinde bulunur ama isimle aranamaz.
- **TC kimlik köprüsü boşken TCKN taşıyan dosyalar otomatiğe geçmez.**
  Ölçülen etki büyüktür: arabuluculuk dosyası **%0**, sağlık listesi **%28**
  otomasyon oranında kalıyor ve 59 satır gereksiz yere incelemeye düşüyor.
  Dosyalarda TC kimlik **var**, eksik olan `veri/tckn_sicil.csv` köprüsüdür.
  Bu köprü İK sisteminden bir kez doldurulursa iki dosya da 0,99 güvenle
  otomatiğe geçer. Projedeki tek en yüksek getirili iyileştirme budur ve kod
  değişikliği gerektirmez.
- **`aile` kuralı ada da takılabiliyor.** Kural soyadı üzerinden çalışır ama
  ölçümde bir vaka adı yakaladı: `SERKAN KOCAK` -> `Kocak Cem` (0,45).
  Güven eşiğinin çok altında kaldığı için yanlış mahsuplaşma üretmedi, satır
  `Eşleşmedi`'ye düştü. Yine de kuralın kesinliği soyad konumunun doğru
  belirlenmesine bağlıdır; iki uçlu isimlerde zayıflar.
- **İşten ayrılmış kişinin adaşı ayırt edilemez.** Ölçümde bulunan tek gerçek
  otomasyon hatası budur (#131 TURAN MEHMET): alias doğru sicile gidiyor ama
  o sicil Nisan 2026'da çıkmış; fatura Temmuz'da yeni işe girmiş **aynı adlı
  başka birine** ait. Sistem bunu çözemez, ama iki uyarı üretip satırı
  incelemeye gönderir — yani hata sessizce geçmez.
- **Doğruluk ölçümünün karşılaştırılabilir örneklemi küçüktür.** Elle
  dağıtılmış dosyanın 134 satırının yalnızca 15'i hem otomasyonla
  karşılaştırılabilir hem de gerçek proje bilgisi taşır. %86,7 doğruluk bu 15
  satır üzerinden hesaplanmıştır. Birkaç ay daha veri biriktikçe bu ölçüm
  güçlenecektir.
- **Elle dağıtılmış dosya doğruluk referansı olarak kullanılamaz.** Yukarıda
  ölçüldüğü gibi o kolon tüzel kişi/ödeyen sorusunu cevaplıyor, proje sorusunu
  değil. İki çıktının farklı olması otomasyonun hatalı olduğu anlamına gelmez;
  karşılaştırma yaparken `testler/dogruluk_olc.py` içindeki `bilgi` okumasına
  bakın.

## Veri gizliliği

Kişisel veri repoya girmez. `.gitignore` şunları dışarıda tutar:

- `ornek_veri/` — personel ana verisi ve gerçek faturalar
- `cikti/` — üretilen Excel dosyaları
- `veri/*.csv` — öğrenen defterler (TC kimlik ve ad soyad içerir)
- `*.pkl` — personel önbelleği

Tek istisna `veri/masraf_merkezi_haritasi.csv`; o bir yapılandırma dosyasıdır,
kişisel veri içermez ve repoda tutulur.

Bütün işlem sizin bilgisayarınızda olur. Hiçbir veri dışarıya gönderilmez,
uygulama internet bağlantısı olmadan çalışır. Öğrenen defterleri yedeklemek
isterseniz şirket içi bir paylaşıma kopyalayın, genel bir depoya koymayın.

## Yeni ay geldiğinde ne yapmalı

1. İK'dan gelen güncel `2025_2026_giris_cikis.xlsx` dosyasını eskisinin
   üzerine yazın. Önbellek dosya boyutu ve değişiklik tarihine bakar, kendini
   yeniler; elle silmenize gerek yoktur.
2. *Ayarlar* sekmesini açıp dönem sayısının arttığını doğrulayın.
3. Ayın fatura dosyalarını *Fatura İşle* sekmesine bırakın ve işleyin.
4. `İncele` sayfasındaki satırları çözün. Her çözüm bir sonraki ay için
   `aliases.csv`'ye yazılır, yani inceleme kuyruğu her ay biraz daha kısalır.
5. Yeni bir görev yeri / proje açıldıysa *Masraf Merkezi Haritası* sekmesinden
   finans kodunu girin. Haritada olmayan görev yerleri uyarı üretir ve görev
   yeri adı olduğu gibi çıktıya yazılır.
6. Yeni bir fatura formatı gelirse önce genel okuyucu denenir. Kolonlar
   tanınmazsa `masraf/okuyucular/` altına o aile için bir okuyucu eklenmelidir.

## Testler

```bash
python3 -m unittest discover -s testler -v
```

Son durum: **147 test, hepsi geçiyor** (yaklaşık 22 saniye). Testler standart
kütüphaneyle yazılmıştır, ek bir test paketi gerekmez.
`ornek_veri/` dizini repoda olmadığı için veri gerektiren testler o dizin
yoksa atlanır (`skipped`); metin normalizasyon testleri her ortamda çalışır.
`testler/test_eslestirici.py` içindeki altın örnekler gerçek Temmuz 2026
verisinden elle doğrulanmış vakalardır ve boş bir öğrenme defteriyle
çalışır — yani her biri sıfırdan kurulan bir sistemde de geçmelidir.
