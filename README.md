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
130'dan 20-30'a indirmek** ve her kararın nedenini görünür kılmak.

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
            +--> transliterasyon varyantı tutuyor mu? --> EŞLEŞTİ (0,88)
            |        (IYLMAZ GEKHAN -> YILMAZ GOKHAN)
            +--> kesik isim öneki tutuyor mu?     --> EŞLEŞTİ  (0,85)
            +--> ek kişi defterinde mi?           --> İNCELE   (0,70)
            +--> bulanık benzerlik yeterli mi?    --> EŞLEŞTİ  (0,79-0,90)
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
| `aile` | 0,60 | Soyadı aynı dosyada kesin eşleşen bir çalışanla aynı; eş/çocuk olabilir |
| çoklu aday | ≤ 0,58 | Aynı isimde birden fazla çalışan var; sicil **doldurulmaz** |
| `yok` | 0,00 | Hiçbir kademe tutmadı |

Karar eşikleri: **0,80 ve üstü** uyarısız satırlar `Sonuç` sayfasında otomatik
kabul edilir. Altındakiler `İncele` sayfasına, hiç eşleşmeyenler `Eşleşmedi`
sayfasına gider. Eşikler *Ayarlar* sekmesinden değiştirilebilir.

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
  beslenmelidir. Ölçüm: Temmuz 2026 dosyasında eşleşmeyen 44 satırın önemli bir
  kısmı tam olarak bu gruptu.
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

Testler standart kütüphaneyle yazılmıştır, ek bir test paketi gerekmez.
`ornek_veri/` dizini repoda olmadığı için veri gerektiren testler o dizin
yoksa atlanır (`skipped`); metin normalizasyon testleri her ortamda çalışır.
`testler/test_eslestirici.py` içindeki altın örnekler gerçek Temmuz 2026
verisinden elle doğrulanmış vakalardır ve boş bir öğrenme defteriyle
çalışır — yani her biri sıfırdan kurulan bir sistemde de geçmelidir.
