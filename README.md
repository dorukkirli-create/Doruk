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

Ölçülen durum (Temmuz 2026, iki Outlook maili, 405 satır; 1C listesi ve
öğrenen defterler açık): 225 satır otomatik dağıtılıyor (%55,6), 166 satır
gerekçesiyle incelemeye düşüyor (%41,0), 14 satırda kişi bulunamıyor (%3,5) ve
tutarı görünür biçimde `(DAGITILAMAYAN)` satırında kalıyor. Yalnız seyahat
dosyası ve 1C listesi olmadan ölçüldüğünde 134 satırın 82'si otomatik, 52'si
insana kalır; o 52'nin 20'si grup şirketi personelidir ve 1C listesi olmadan
**hiçbir zaman** otomatik dağıtılamaz. Sayılar ve nasıl üretildikleri aşağıda
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
            +--> harici_kisiler.csv'de var mı?    --> İNCELE   (0,95) "çalışan değil"
            +--> tam isim tek kişide mi?          --> EŞLEŞTİ  (0,95)
            +--> 1C personel listesinde birebir mi? --> İNCELE  (0,90 tek / 0,55 çok aday)
            |        (bilinen kimlik tahminden önce; grup şirketi personeli,
            |         dönem doğrulanamaz; sicil ana veride de varsa 0,92 tam isim)
            +--> isim alt kümesi tek kişide mi?   --> EŞLEŞTİ  (0,90; soyad tutmuyorsa 0,72 İNCELE)
            +--> bitişik ad açılıyor mu?          --> EŞLEŞTİ  (0,92)
            |        (AHMETCAN -> MUSTAFA KEMAL)
            +--> transliterasyon varyantı tutuyor mu? --> İNCELE  (0,88)
            |        (IYLMAZ GEKHAN -> YILMAZ GOKHAN)
            +--> kesik isim öneki tutuyor mu?     --> İNCELE   (0,85)
            +--> ek kişi defterinde mi?           --> İNCELE   (0,70)
            +--> bulanık benzerlik yeterli mi?    --> İNCELE   (en çok 0,89; asla otomatik değil)
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
  [6] Mahsuplaşma: yineleme eleme, paylaşım, kuruşuna kadar mutabakat
            |
            v
  [7] Excel çıktısı (9 sayfa): Ozet | Mahsuplasma | Kontrol | Dosyalar | Sirket Kirilimi
                               | Harita Onerileri | Sonuc | Incele | Eslesmedi
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

İki kullanım yolu vardır; ikisi aynı çekirdek kodu (`masraf/`) kullanır.

**1) Finans ekibi için masaüstü paketi (internet ve Python kurulumu gerekmez).**
`paketle/` altındaki tarif ile üretilen `Otomasyon\` klasörü: faturalar
`1_FATURALAR`'a atılır, `CALISTIR.bat` çift tıklanır, Excel `2_EXCEL_CIKTI`'ya
çıkar, işlenen faturalar `3_ISLENENLER`'e taşınır. Kurulum ve kullanım
`paketle/windows/OKU_BENI.txt`, paketleme tarifi `paketle/BENIOKU.md`. Bu yolda
arayüz yoktur; öğrenme otomatik defter beslemesiyle ve `veri/*.csv`
dosyalarının elle düzenlenmesiyle olur.

**2) Geliştirici / analist için Streamlit arayüzü (ilk kurulumda internet
gerekir).** Aşağıdaki adımlar bu yol içindir.

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
4. **Mahsuplaşmayı okuyun.** *Mahsuplaşma* sekmesi muhasebeye gidecek tabloyu
   gösterir: her fatura için hangi projeye ne kadar yazılacağı. En üstte
   mutabakat vardır; yeşilse para kaybolmamış demektir.
5. **İnceleyin.** *İnceleme* sekmesinde düşük güvenli satırlar tek tek gelir.
   Her satırın yanında sistemin **neden** o sonuca vardığı Türkçe yazar ve
   adaylar listelenir. Doğru kişiyi seçtiğinizde sistem bunu öğrenir.
6. **Excel'i indirin.** Dokuz sayfalı çıktı, RHI kurumsal kimliğinde:
   `Ozet` (kapak: tutarlar, mutabakat durumu, şirket kırılımı, dikkat notları),
   `Mahsuplasma` (muhasebeye giden dağıtım tablosu), `Kontrol` (fatura bazında
   mutabakat), `Dosyalar` (dosya envanteri: hangi dosya ve ek okundu, hangisi
   atlandı, neden; satır ve tutar), `Sirket Kirilimi` (tüzel kişi üstte, projeleri altında),
   `Harita Onerileri` (tanımsız görev yerleri için hazır satırlar), `Sonuc`
   (tüm satırlar, evrak numarasıyla), `Incele` (elle bakılacaklar), `Eslesmedi`
   (kişi bulunamayanlar).
7. **Klasör kendini toplar.** Masaüstü paketinde işlenen faturalar
   `3_ISLENENLER\<tarih_saat>\` altına taşınır, her çalıştırma
   `CALISTIRMA_GECMISI.txt` dosyasına kaydedilir. Gelecek ay aynı dosyalar
   yanlışlıkla tekrar işlenmez.

## Çıktı nasıl görünür

Nihai çıktı satır dökümü değil, **dağıtım tablosudur**. Her fatura için altında
hangi projeye ne kadar yazılacağı yazar:

| Fatura | Masraf Merkezi | Şirket | Gider Tipi | Tutar | Fatura Payı | Satır | Kişi | Durum |
|---|---|---|---|---|---|---|---|---|
| ENERGO TEMMUZ.xls | GPP | UST LUGA | Bilet | 22.327,76 | %45,6 | 62 | 54 | 2 eşleşmeyen satır |
| ENERGO TEMMUZ.xls | HQ-MOSCOW | RHI | Bilet | 4.428,18 | %9,1 | 12 | 5 | 7 satır incelenecek |
| ENERGO TEMMUZ.xls | HQ-MOSCOW | RHI | Otel | 2.397,06 | %4,9 | 7 | 4 | |
| ENERGO TEMMUZ.xls | (DAGITILAMAYAN) | | Diğer | 2.142,05 | %4,4 | 2 | 0 | MASRAF MERKEZI YOK |

Tablonun üç kuralı vardır.

**1. Her fatura kendi içinde kapanır.** Okunan tutar = yinelenen + dağıtılan +
dağıtılamayan. Kuruşuna kadar. Oransal bölmede kalan artık, faturanın en büyük
satırına eklenir; `Kontrol` sayfasındaki **Fark** sütunu sıfır olmak zorundadır.
Sıfır değilse tablo muhasebeye gönderilmemelidir.

**2. Dağıtılamayan tutar görünür kalır.** Kişisi ya da masraf merkezi
bulunamayan tutar sessizce düşürülmez; `(DAGITILAMAYAN)` satırı olarak durur.
Toplam her zaman faturaya eşittir.

**3. Yinelenen işlemler bir kez sayılır.** Aynı mail iki dosya taşıyabilir:
acentenin ham cari dökümü ve aynı işlemlerin elle dağıtılmış hali. İkisi de
okunursa para **çift sayılır**. Temmuz 2026 mailinde ölçüldü:

| Dosya | Satır | Tutar |
|---|---|---|
| ENERGO TEMMUZ.xls (ham döküm) | 134 | 48.946,59 USD |
| YUZYIL TEMMUZ.xlsx (elle dağıtılmış) | 134 | 48.978,59 USD |

Bu ikisi **aynı** işlemlerdir. Naif toplama 97.925,18 USD verir; gerçek rakam
yarısıdır. Tek fark 14.07.2026 tarihli bir kalemdir: ham dökümde -16,00 USD
(iade), elle dağıtılmış halde +16,00 USD. Aradaki 32,00 USD tam olarak budur ve
işaret çelişkisi olarak ayrıca uyarı verilir.

Eşleştirme kaba bir kova (belge tarihi + mutlak tutar + para birimi) içinde
yapılır, sonra kova içinde isimler eşleştirilir. İsim anahtar olarak
kullanılmaz çünkü iki dosya aynı kişiyi farklı yazar: ham döküm
`DEMIRALP AHMETCAN`, elle dağıtılmış hal `AHMET CAN DEMIRALP`, bazen de
kırpılmış (`DEMIRALP AHMETCA`) ya da yanlış (`KARATAS ANIL` /
`ALI YALCINKAYA`). Kişi adı olmayan kurumsal kalemler (cenaze çelengi,
toplantı organizasyonu) da bu sayede yakalanır.

Elenen kayıt atılmadan önce taşıdığı **dağıtım talimatı** tutulan kayda
aktarılır. İki dosya birbirini tamamlar: ham döküm tutarı doğru taşır, elle
dağıtılmış hal insanın aldığı kararı taşır. Örnek: `RHI 1/3 - RENSTROYDETAL
2/3` paylaşımı yalnızca elle dağıtılmış dosyada vardır; devralınmasa kaybolurdu.

Paylaşım etiketleri pratikte **tüzel kişi** adıdır, proje değil. Bu yüzden
proje aynı kalır, tutar şirketler arasında bölünür. Etiket masraf merkezi
haritasında gerçekten bir projeye karşılık geliyorsa o zaman proje de bölünür.

Kişi listeleri (katılımcı listesi, sağlık kontrol listesi) dağılıma girmez;
bunlar fatura değil kütüktür ve tutar taşımazlar. Temmuz 2026 mailinde 106
satır bu gruptadır. Aynı gruba tedarikçinin fatura başına gönderdiği
`ASS... Fatura Detayı.xlsx` katılımcı listeleri de girer: tutar kolonu yoktur,
tutar yansıtma dosyasındadır. Bu listeler atılmaz; kişileri aynı fatura
numaralı yansıtma satırlarıyla çapraz kontrol edilir ve fark varsa
(yansıtmada eksik ya da fazla kişi) `Kontrol` sayfasında ve kapakta uyarı
olarak görünür.

Arabuluculuk dosyasında kişi başı tutar şirket toplamının kişi sayısına
bölünmesiyle bulunur. Bölme kuruşta yapılır ve artık kuruşlar ilk kişilere
birer birer eklenir; böylece kişi tutarlarının toplamı faturanın kendi beyan
ettiği toplama kuruşuna kadar eşittir (Temmuz 2026: 1.943,74 USD, fark 0,00).

## Desteklenen dosya tipleri

| Kaynak | Tanıma ipucu | Kimlik anahtarı | Masraf merkezi |
|---|---|---|---|
| Antik / Yüzyıl ham cari hareket dökümü (`.xls`) | "Cari Hareket Dökümü Detay" başlığı, `İşlem`/`Evrak No`/`Borç` kolonları | Kişi adı açıklama metnine gömülü | Yok, eşleştirmeden gelir |
| Yüzyıl elle dağıtılmış (`.xlsx`) | `S.NO`, `AÇIKLAMA`, `ŞANTİYESİ` kolonları | Kişi adı açıklamada | Var (`ŞANTİYESİ`) — doğruluk referansı |
| Energo assessment yansıtma (`.xlsx`) | `Fatura Detay` + `Kişi Listesi` sayfaları, `Katılımcı` kolonu | Ad Soyad | Yok |
| Energo assessment fatura detay listesi (`ASS... Fatura Detayı.xlsx`) | `Katılımcı` + `Paket` kolonları var, tutar kolonu yok | Ad Soyad | Yok; kütük sayılır, yansıtma ile çapraz kontrol edilir |
| Energo arabuluculuk (`.xlsx`) | `PERSONEL T.C.`, `PROJE` kolonları | **TC kimlik no** | Var (`PROJE`) |
| Sağlık kontrol listesi (`.xlsx`) | `BORDROLU LİSTE` sayfası, `TCKN` + `ŞANTİYE` | **TC kimlik no** + doğum tarihi | Var (`ŞANTİYE`) |
| Koç Üniversitesi katılımcı listesi (`.xlsx`) | `ID`, `Ad Soyad`, `Katılım Tarihi` | **Sicil numarası** (`ID` kolonu) | Yok |
| Outlook e-postası (`.msg`) | Dosya uzantısı | Ekteki dosyaya göre | Ekteki dosyaya göre |
| Kişi kütüğü (`referans_liste`) | 200+ satır ve hiçbir satırda tutar yok (sigorta listesi gibi) | Defter beslemesinde kullanılır | Dağılıma girmez; hangi dosyanın kütük sayıldığı uyarıda yazar |
| Tanınmayan tablo (`.xlsx`, `.csv`) | Genel okuyucu, kolon adlarından çıkarım | Bulunabilene göre | Varsa okunur |

En kolay eşleşen kaynaklar sicil veya TC kimlik taşıyanlardır. Seyahat
faturaları en zorudur: kimlik alanı yoktur, sadece serbest metinde isim vardır.

## Güven skorları ne demek

| Yöntem | Güven | Ne zaman oluşur |
|---|---|---|
| `sicil` | 1,00 / 0,75 | Kaynak dosyada sicil numarası var (Koç katılımcı listesi gibi). Satırdaki ad ile personel adının hiçbir kelimesi tutmuyorsa 0,75 (yanlış satıra yazılmış sicil), İNCELE |
| `tckn` | 0,99 / 0,75 | TC kimlik `veri/tckn_sicil.csv` köprüsünde tek bir sicile bağlanıyor; ad hiç tutmuyorsa 0,75, İNCELE |
| `harici` | 0,95 | Kişi `veri/harici_kisiler.csv` dış kişi defterinde (danışman, konuşmacı). Çalışan olmadığı için her zaman uyarılı, İNCELE'ye düşer; masraf merkezi defterden gelir |
| `yardimci_defter` | 0,90 / 0,55 | Kişi 1C personel listesinde birebir (grup şirketi). Tam isimden hemen sonra, tahmin kademelerinden ÖNCE denenir. Tek aday 0,90, çok aday 0,55; liste tek tarihli olduğu için dönem doğrulanamaz, İNCELE. Sicil ana veride de varsa 0,92 `tam_isim` olur |
| `alias` | 0,98 / 0,75 | Bu ismi daha önce siz elle onaylamışsınız (`veri/aliases.csv`). Ana veride aynı isimli başka çalışan da varsa 0,75 ve adaylar listelenir, İNCELE |
| `tam_isim` | 0,95 | Normalize isim personel verisinde **tek** kişiye denk geliyor |
| `tam_isim` (bitişik ad) | 0,92 | `AHMETCAN` sözlükle `MUSTAFA KEMAL` olarak açıldı, sonuç tek kişi |
| `alt_kume` | 0,90 / 0,72 | Fatura ismi personel isminin alt kümesi, tek aday (ikinci ad/patronimik eksik). Fatura kelimelerinden biri personelin SOYADI değilse 0,72, İNCELE |
| `transliterasyon` | 0,88 | Rusça transliterasyon geri çevrildi (`GEKHAN` -> `GOKHAN`), tek aday |
| `prefix` | 0,85 | İsim bilet sisteminde kesilmiş, önek tek kişiye uyuyor |
| `alt_kume` (ters) | 0,72 / 0,50 | Personel adının tamamı fatura metninin içinde (ayıklama artığı). Fazladan kalan kelime bir AD ise (üçüncü adı olan başka kişi olabilir) sicil doldurulmaz, 0,50 |
| `ek_defter` | 0,70 | Kişi ana veride yok, yardımcı listelerden (sağlık listesi vb.) bulundu |
| `bulanik` | 0,79-0,89 | Yazım hatası toleranslı benzerlik (rapidfuzz ≥ 88 puan), ikinci adayla arada en az 6 puan fark. Tavan 0,89: asla otomatik kabul edilmez. Tek kelimelik/placeholder personel adları ve alt küme durumları havuza girmez; iki kelimeli adlarda farklı kalan kelime çifti yakın değilse (ALI/ANIL) sicil doldurulmaz |
| `aile` | 0,30-0,60 | Soyadı aynı çalıştırmadaki bir FATURADA kesin eşleşen bir çalışanla aynı; eş/çocuk olabilir. Kanıt yalnızca kesin yöntemlerden (sicil, TC kimlik, alias, tek kişilik tam isim) ve fatura tipi kaynaklardan gelir; sağlık/katılımcı listeleri kanıt üretmez. Kanıtın gücüne göre dört kademe |
| çoklu aday | ≤ 0,58 | Aynı isimde birden fazla çalışan var; sicil **doldurulmaz** |
| `yok` | 0,00 | Hiçbir kademe tutmadı |

Karar eşikleri (`masraf/masraf_merkezi.py`): **0,90 ve üstü** güvene sahip
*ve* hiç uyarı taşımayan satırlar `Sonuç` sayfasında otomatik kabul edilir
(`GUVEN_ESIGI = 0.90`). **0,50'nin altındakiler** eşleşmemiş sayılır ve
`Eşleşmedi` sayfasına gider (`ALT_ESIK = 0.50`). Arada kalan her şey `İncele`
sayfasına düşer. Eşikler *Ayarlar* sekmesinden değiştirilebilir.

Buradan çıkan sonuç önemlidir: `transliterasyon` (0,88), `prefix` (0,85),
`bulanik` (en çok 0,89), `ek_defter` (0,70) ve `aile` (0,60) kademelerinin
**hiçbiri kendi başına otomatik kabul edilmez**; hepsi insan onayına gider.
Satırda TC kimlik varken tahmin kademesi Türkiye vatandaşı olmayan birine
giderse güven 0,75'e çekilir ve uyarı yazılır.
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

#### 1C ikincil personel defteri açıkken

Yukarıdaki tablo **yalnızca ana veriyle** ölçülmüştür. Uygulama, yanında bir
1C personel listesi bulduğunda onu ikincil defter olarak da kullanır
(grup şirketlerini kapsar, ölçülen katkı: 17.517 isimli kayıttan 5.234'ü ana
veride yoktur). Aynı örneklem, 1C defteri açıkken:

| Dosya | Satır | OTOMATİK | İNCELE | EŞLEŞMEDİ | Otomasyon |
|---|---:|---:|---:|---:|---:|
| `ANTIK_CARI_TEMMUZ_2026.xls` | 134 | 82 | 45 | 7 | %61,2 |
| `YUZYIL_..._ELLE_DAGITILMIS.xlsx` | 134 | 79 | 48 | 7 | %59,0 |
| `ASSESSMENT_YANSITMA_2026_05_06.xlsx` | 6 | 4 | 2 | 0 | %66,7 |
| `ARABULUCULUK_2026_06_07.xlsx` | 25 | 0 | 25 | 0 | %0,0 |
| `SAGLIK_KONTROL_LISTE.xlsx` | 50 | 14 | 36 | 0 | %28,0 |
| `KOC_UNI_KATILIMCI_LISTESI.xlsx` | 50 | 46 | 4 | 0 | %92,0 |
| **TOPLAM** | **399** | **225** | **160** | **14** | **%56,4** |

Çözülen oran **%89,2 -> %96,5**, eşleşmeyen satır **43 -> 14**. Otomasyon oranı
**değişmez (%56,4)**, çünkü 1C listesi tek bir tarihe ait durum fotoğrafıdır,
aylık dönem serisi değildir; bu defterden gelen her kayıt "dönem
doğrulanamadı" uyarısı taşır ve bilerek incelemeye gönderilir. Yani 1C defteri
**otomatik dağıtımı artırmaz, "hiç bulunamadı" sayısını azaltır** — kullanıcıya
boş satır yerine doğrulanacak bir aday verir.

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

### Eşleşmeyen 43 satırın nedeni (1C defteri kapalıyken)

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

- **#131, adaş vakası** (ESB-LED, 31.07). Bulunan sicil Amursky'de çalışmış ama
  **06.04.2026'da çıkmış**. Aynı günün aynı partisindeki iki kişi için hem elle
  hem otomasyon "GPP" diyor. Bu neredeyse kesinlikle yeni işe girmiş **aynı adlı
  başka bir kişi**. Elle dosya doğru, otomasyon yanlış; ama sistem iki uyarıyla
  ("gider ayında personel kaydı yok", "belge tarihinden önce ayrılmış") tam
  olarak doğru yeri işaret etti ve satırı otomatik kabul etmedi.
- **#69, ad-soyad karışması.** `aile` kuralı soyada değil **ada** takılıp aynı adı
  soyadı olarak taşıyan bir çalışana 0,45 güvenle bağlandı. Eşik altında kaldığı
  için `Eşleşmedi` sayfasına düştü, yanlış mahsuplaşma üretmedi.

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
(seyahat dökümündeki `TK4093099626 DEMIRALP/AHMETCAN MR IST-CDG BILET BEDELI`
gibi) kolon sözlüğü yetmez, o kalıp için kod yazmak gerekir. Yılda bir iki kez
karşılaşılacak bir durumdur.

---

## Bilinen kısıtlar

Bunları bilerek kullanın; araç bunları gizlemez, çıktıda uyarı olarak gösterir.

- **Personel ana verisinde TC kimlik ve pasaport numarası yoktur.** Sadece
  sicil, ad soyad ve doğum tarihi kimlik alanı olarak bulunur. TC kimlik ile
  eşleşme yapabilmek için `veri/tckn_sicil.csv` köprüsü elle veya İK
  sisteminden doldurulabilir. Ancak köprü artık **otomatik de türetilir**:
  sağlık listesindeki ad soyad ve doğum tarihi personel verisiyle birleştirilir
  ve tek adaya inen kayıtlar köprüye yazılır. Ölçüm: 50 kişinin 27'si tek
  çalıştırmada bağlandı.
- **Ana personel dosyası yalnızca RHI ve UST LUGA tüzel kişilerini kapsar.**
  Renservis, Renstroydetal, One Tower, Top Tower, RC Peter, RC Moskova
  personeli bu dosyada **yoktur**. Bu kısıt 1C personel listesi eklenerek
  kapatılır (bkz. "İki personel dosyası kullanın"); 1C listesi olmadan bu
  kişiler eşleşmedi olarak gelir.
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
  Hint uyruklu personelde %10,9'dur (çok yaygın Hint ad-soyad çiftleri).
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
  ölçümde bir vaka adı yakaladı: kişinin ADI, başka bir çalışanın SOYADI (0,45).
  Güven eşiğinin çok altında kaldığı için yanlış mahsuplaşma üretmedi, satır
  `Eşleşmedi`'ye düştü. Yine de kuralın kesinliği soyad konumunun doğru
  belirlenmesine bağlıdır; iki uçlu isimlerde zayıflar.
- **İşten ayrılmış kişinin adaşı ayırt edilemez.** Ölçümde bulunan tek gerçek
  otomasyon hatası budur (#131, adaş vakası): alias doğru sicile gidiyor ama
  o sicil Nisan 2026'da çıkmış; fatura Temmuz'da yeni işe girmiş **aynı adlı
  başka birine** ait. Sistem bunu çözemez, ama iki uyarı üretip satırı
  incelemeye gönderir — yani hata sessizce geçmez.
- **Doğruluk ölçümünün karşılaştırılabilir örneklemi küçüktür.** Elle
  dağıtılmış dosyanın 134 satırının yalnızca 15'i hem otomasyonla
  karşılaştırılabilir hem de gerçek proje bilgisi taşır. %86,7 doğruluk bu 15
  satır üzerinden hesaplanmıştır. Birkaç ay daha veri biriktikçe bu ölçüm
  güçlenecektir.
- **Ölçüm betikleri 1C ikincil defterini kullanmıyor.**
  `testler/kapsam_olc.py` ve `testler/dogruluk_olc.py` yalnızca ana veriyle
  çalışır; uygulamanın kendisi ise yanında bir 1C listesi bulduğunda onu
  otomatik kullanır. Bu yüzden betiklerin bastığı eşleşmeyen sayısı (43),
  uygulamanın gerçek davranışına (14) göre **kötümserdir**. Yukarıdaki
  "1C ikincil personel defteri açıkken" tablosu farkı gösterir. Betiklere bir
  `--yardimci` seçeneği eklenmesi bekleyen bir iştir; ölçümlerin ikisi de
  düzeltilene kadar alt sınır olarak okunmalıdır.
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

Son durum: **250 test, hepsi geçiyor** (yaklaşık 50 saniye). Testler standart
kütüphaneyle yazılmıştır, ek bir test paketi gerekmez.
`ornek_veri/` dizini repoda olmadığı için veri gerektiren testler o dizin
yoksa atlanır (`skipped`); metin normalizasyon testleri her ortamda çalışır.
`testler/test_eslestirici.py` içindeki altın örnekler gerçek Temmuz 2026
verisinden elle doğrulanmış vakalardır ve boş bir öğrenme defteriyle
çalışır — yani her biri sıfırdan kurulan bir sistemde de geçmelidir.
