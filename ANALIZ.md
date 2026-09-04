# Masraf Merkezi Tespiti: Veri Analizi

Bu doküman, finans ekibinden gelen örnek ham dosyalar ile 2025-2026 giriş çıkış verisinin
incelenmesi sonucu ölçülen bulguları içerir. Buradaki her sayı gerçek veriden hesaplanmıştır,
tahmin değildir.

İnceleme tarihi: Eylül 2026
İnceleme kapsamı: Mustafa Demirel'in ilettiği iki örnek mail zinciri ve ekleri, 2025-2026 giriş çıkış dosyası.

---

## 1. Kısa cevap

Evet, yapılabilir. Ama tek bir dosya ile değil, katmanlı bir kimlik çözümü ile.

Ölçülen sonuç: Temmuz 2026 seyahat faturasındaki 130 kişili satırın 63'ü ham isimle doğrudan,
22'si kısmi eşleşme ile personel veritabanında bulundu. Yani naif bir eşleştirme ile yüzde 66.
Kalan yüzde 34'ün neden bulunamadığını tek tek açtım ve bunların çoğu çözülebilir teknik
problemler, veri eksikliği değil. Detay bölüm 4'te.

Kritik kısıt: paylaştığınız hiçbir dosyada pasaport numarası yok. TC kimlik numarası ise sadece
yardımcı listelerde var, personel ana verisinde yok. Bu yüzden şu an kimlik eşleştirmesi
isim üzerinden yürümek zorunda. Bunu değiştirecek tek hamle bölüm 7'de.

---

## 2. Veri kaynaklarının haritası

Finans ekibine gelen dosyalar tek bir formatta değil. En az altı farklı şablon var ve her biri
farklı bir kimlik anahtarı taşıyor. Bu, otomasyonun ana zorluğu.

| Kaynak | Kimlik anahtarı | Masraf merkezi var mı | Kişi adı formatı |
|---|---|---|---|
| Antik / Yüzyıl Travel ham cari döküm (.xls) | yok, sadece isim | hayır | serbest metin içine gömülü |
| Yüzyıl elle dağıtılmış (.xlsx) | yok, sadece isim | evet, elle yazılmış | ayrı kolonda |
| Energo Assessment (AS Ölçme) | yok, sadece isim | şirket düzeyinde (RHI/RSD) | Ad Soyad |
| Energo Arabuluculuk | TC kimlik no | evet, proje adı yazılı | Ad Soyad |
| Sağlık Kontrol Listesi | TC kimlik no | evet, şantiye yazılı | Ad Soyad |
| Koç Üniversitesi katılımcı listesi | sicil numarası | hayır | Soyad Ad |
| Ferdi Kaza sigorta listesi | sicil numarası | evet, doğrudan masraf merkezi kolonu | Soyad Ad |

En kolay kaynak Koç Üniversitesi ve Ferdi Kaza listeleri, çünkü sicil numarası taşıyorlar.
En zor kaynak seyahat acentesinin ham dökümü, çünkü kişi adı fatura açıklamasının içinde gömülü.

### Seyahat acentesi ham dosyasında isim nasıl duruyor

Dosya bir cari hareket dökümü. Masraf merkezi kolonu hiç yok. Kişi adı açıklama metninin içinde
ve dört farklı kalıpta geçiyor:

```
TK4093099626 OZAKAY/MUSTAFAKEMAL MR  IST-CDG BILET BEDELI
PC2255749381 TEMIR MEHMET\M  KYA-SAW-LED BILET BEDELI
CIHAN BALABAN GRAND PLAZA HOTEL HANOI [11.07.2026] - [13.07.2026]  (2) KONAKLAMA YURTDISI
EKSTRA BAGAJ UCRETI ISA MUCAHIT SAHIN TARAFINDAN TASINDI
```

Bilet satırlarında soyad önce geliyor, otel satırlarında ad önce geliyor. Otel adı isme yapışık.
Bunların hepsi ayrıştırılabilir, ama her kalıp için ayrı kural gerekiyor.

---

## 3. Personel ana verisinin yapısı ve kalitesi

2025_2026_giris_cikis.xlsx aylık snapshot mantığında çalışıyor. Her kişi her dönem için bir satır.

| Ölçüm | Değer |
|---|---|
| Toplam satır | 153.675 |
| Dönem aralığı | Kasım 2025 ile Temmuz 2026 arası, 9 ay |
| Benzersiz sicil | 28.767 |
| İsmi dolu benzersiz sicil | 24.307 |
| Adı Soyadı boş satır | 44.482 |
| Çıkış kaydı olan kişi | 12.227 |

### Bulunan veri kalitesi sorunları

Bunlar otomasyonu sessizce bozacak türden sorunlar, önceden bilmekte fayda var.

1. **Sicil kolonu karışık tipte.** 145.261 satır metin, 8.414 satır sayı olarak duruyor.
   Ayrıca C92620 ve D33935 gibi harfli siciller var. Tip dönüşümü yapılmazsa birleştirme
   sessizce başarısız olur ve kimse fark etmez. İlk analizimde bu yüzden isim çakışma oranını
   yüzde 39 ölçtüm, düzelttikten sonra gerçek oran yüzde 3,6 çıktı.

2. **44.482 satırda isim boş.** Bunlar bordrosuz taşeron kayıtları, sicilleri 900001 gibi
   sahte numaralar. İsim indeksine girerlerse yanlış eşleşme üretirler.

3. **Kategori kolonunda iki yazım var.** Türkçe karakterli "Çıkış" 10.544 satır, ASCII "Cikis"
   2.246 satır. Filtreleme yaparken ikisini de yakalamak gerekiyor.

### İyi haber: çıkış kayıtları temiz

Çıkış almış kişilerin projesini sormuştunuz. Bu tarafta veri sağlam:

| Ölçüm | Değer |
|---|---|
| Çıkış satırı | 12.790 |
| Çıkış tarihi dolu | yüzde 100 |
| Görev yeri dolu | yüzde 100 |
| Adı soyadı dolu | yüzde 100 |
| Resmi çıkış sebebi dolu | 8.124 satır |

Yani bir kişi fatura tarihinden önce çıkmışsa, son çalıştığı projeyi kesin olarak bulabiliyoruz.
Otomasyon bu durumu ayrı bir uyarı ile işaretliyor, çünkü çıkmış birinin masrafı genelde
kontrol edilmesi gereken bir durumdur.

### İsim çakışması gerçekte sorun değil

Endişe edilecek konu isim benzersizliğiydi. Ölçtüm:

| Popülasyon | Kişi | İsim çakışma oranı |
|---|---|---|
| Tümü | 24.307 | yüzde 3,6 |
| Beyaz yaka | 4.395 | yüzde 2,5 |
| Ekspat | 5.760 | yüzde 7,8 |
| Beyaz yaka ve ekspat | 1.241 | yüzde 1,5 |
| Şirket RHI ve beyaz yaka | 1.118 | yüzde 0,2 |
| Vatandaşlık Hindistan | 3.837 | yüzde 10,9 |

Seyahat, eğitim ve değerlendirme masrafı yapan popülasyon ağırlıklı olarak RHI beyaz yaka.
Orada çakışma binde iki. Yani isimle eşleştirme bu iş için yeterince güvenli.

Çakışmanın yoğunlaştığı yer Hindistan vatandaşı mavi yaka personel. Kumar Manoj ismi 14 farklı
sicile ait. Bu kişiler seyahat faturası üretmiyor, ama otomasyon yine de çoklu aday durumunda
otomatik karar vermemeli.

### Masraf merkezi neredeyse hiç değişmiyor, ama yine de tarihe bakmak gerekiyor

Kişilerin sadece yüzde 1,4'ünün görev yeri, yüzde 3,5'inin şirketi dönemler arasında değişiyor.
Küçük bir oran. Ama tam olarak o yüzde 1,4 yanlış mahsuplaşmaya yol açar, ve bunlar genelde
proje değiştiren kıdemli kişiler yani tutarı yüksek masraflar. Bu yüzden otomasyon her satırda
fatura tarihine en yakın dönemin kaydını kullanıyor, en güncel kaydı değil.

---

## 4. Eşleşmeyen 44 satır neden eşleşmedi

Bu bölüm işin özü. Her eşleşmeme sebebini tek tek açtım ve hangisinin çözülebilir olduğunu ayırdım.

### Çözülebilir: aile bireyleri

Şirket hesabından eş ve çocuk bileti alınıyor. Bunlar personel değil, ama masrafı çalışanın
projesine gitmeli.

| Faturadaki isim | Bağlı olduğu çalışan | Sicil | Masraf merkezi |
|---|---|---|---|
| GUNAL DARIA, GUNAL SERAFIMA, GUNAL YENISEY | Gunal Emre | 102084 | RHI Russia Headquarter Moscow |
| ARAS CELENLIGIL, SEVIM CELENLIGIL | Celenligil Onur | 423806 | ALNG2-GBS Project |
| COSKUN ELIF MUALLA GOKCE | Coskun Emre | 512495 | GPP Project |
| ECE OLGA, UGUR OLGA | soyadı eşleşen çalışan | - | - |

Kural: soyadı tek bir çalışan ile eşleşiyorsa aile bireyi olarak işaretle, masraf merkezini
o çalışandan devral, ama güven skorunu düşür ve incelemeye gönder. Otomatik kabul etme.

### Çözülebilir: Rusça transliterasyon

İsimler Rusça belgelerden Latin alfabesine geri çevrilirken bozuluyor.

| Faturadaki yazım | Gerçek kişi | Sicil | Masraf merkezi |
|---|---|---|---|
| YRMAK MEKHMET VEISI | Irmak Mehmet Veysi | 300973 | Regional Management RHI |
| IYLMAZ GEKHAN | Yilmaz Gokhan | 105045 | RHI Russia Headquarter Moscow |

Bu iki satır elle dağıtılmış dosyada RENSERVIS ve RHI olarak işaretlenmiş. Personel verisine göre
Irmak Mehmet Veysi RHI Regional Management'ta. Yani elle yapılan dağıtımda bir hata var gibi
görünüyor. Doğrulanması gereken bir nokta.

### Çözülebilir: kesilmiş ve bitişik isimler

Excel kopyalama sırasında isimler 20 karakterde kesilmiş.

| Faturadaki yazım | Gerçek kişi | Sicil |
|---|---|---|
| OZAKAY/MUSTAFAKEMA | Ozakay Mustafa Kemal | 100003 |
| ALLANAZAROV/ALLANAZA | Allanazarov ailesinden biri | - |
| GRINEVICH/NATALIA | - | - |

Ayrıca MUSTAFAKEMAL gibi bitişik yazılmış çift adlar var. Sözlük tabanlı ayırma ile çözülüyor.

### Çözülebilir: metin gürültüsü

Otel adı, vize metni ve bagaj açıklaması isme yapışık geliyor. Önce temizlik, sonra eşleştirme
gerekiyor.

### Çözülemez: veride gerçekten olmayan kişiler

Bu grup üç alt kategoriye ayrılıyor.

**Grup şirketi personeli.** Personel ana verisi sadece RHI ve UST LUGA tüzel kişilerini kapsıyor.
Renservis, Renstroydetal, RC Peter, RC Moskova, One Tower, Top Tower personeli bu dosyada yok.
Örnek: Tolga Surul (One Tower), Koray Erdur (Top Tower), Vedat Boynuegri (Renservis).

**Dış danışmanlar.** Talip Kerem Kockesen bir kurumsal gelişim koçu, çalışan değil. Faturası
"KURUMSAL GELİŞİM KOÇLUĞU BEDELİ" olarak geliyor. Bu kişilerin ayrı bir dış kişi defterinde
tutulması gerekiyor.

**Henüz sicili olmayan yeni girenler.** Sizin bahsettiğiniz durum tam olarak bu. Ankara Peter ve
Ankara Moskova uçuşları mobilizasyon uçuşları, kişi henüz işe başlamamış.

### Ve burada iyi bir haber var

Bu son grubun büyük kısmı zaten elinizdeki başka bir dosyada duruyor. Eşleşmeyen 43 ismi
Sağlık Kontrol Listesi ile karşılaştırdım:

**15 tanesi, yani yüzde 35'i, sağlık kontrol listesinde TC kimlik numarası ve şantiye bilgisiyle
mevcut.**

| İsim | TC kimlik | Şantiye |
|---|---|---|
| NEVZAT GULER | 63574287888 | Ust Luga Gas Processing Complex GPC |
| GOKHAN GUZEL | 55012637188 | Ust Luga GPP projesi |
| ENIS DONMEZ | 72010082796 | Ust Luga GPP projesi |
| HARUN YILDIZ | 34448323056 | Ust Luga GPP projesi |
| AHMET CELER | 39424783404 | UST LUGA |
| MUSTAFA OZCAN | 34579781688 | LYTKARINO |
| HUSEYIN BAGUC | 11662308052 | Ust Luga Gas Processing Complex GPP |
| OGUZHAN CANKAYA | 10596168220 | Ust Luga Gas Processing Complex GPC |
| FURKAN ERDEM | 59497164320 | Ust Luga Gas Processing Complex GPC |
| HARUN NERGIS | 23260033694 | Ust Luga GPP projesi |
| MEHMET TURAN | 26290978980 | Ust Luga Gas Processing Complex GPC |
| MUHAMMED YILMAZ | 28160415434 | Amur AGPP |
| ISA MUCAHIT SAHIN | 13906007882 | Ust Luga Gas Processing Complex GPC |
| ALI GUNDOGDU | 10405297660 | UST LUGA, Renstroydetal |
| UMUT OZTURK | 39838874224 | Ust Luga GPP projesi, Yaka LLC |

Yani otomasyonun bir "ek kişi defteri" tutması ve bunu sağlık listesi gibi yardımcı kaynaklardan
otomatik beslemesi gerekiyor. Bu tek başına eşleşme oranını ciddi şekilde yukarı çeker.

---

## 5. Masraf merkezi sözlüğü tutarlı

Ferdi Kaza sigorta listesinin Data sayfasında 9.195 satırlık bir kayıt var ve içinde doğrudan
bir MASRAF MERKEZI kolonu bulunuyor. Değerleri kontrol ettim:

```
GPP Project, Ust Luga Fabrication - GPC (RHI), Amursky Gas Processing Plant,
ALNG2-GBS Project, RHI Russia - Headquarter (Moscow), ALNG2-Gydan,
Ust Luga Steel Package, Project Control Office, Udokan (GMK),
Regional Management (RHI), Novosibirsk Technical Office
```

Bunlar giriş çıkış dosyasındaki Görev Yeri kolonunun değerleriyle birebir aynı sözlüğü kullanıyor.
Bu önemli, çünkü şu sonuca varıyoruz:

**Personel dosyasındaki Görev Yeri kolonu zaten masraf merkezidir.** Ayrı bir eşleme tablosu
kurmaya gerek yok, sadece bu değerlerin finans tarafındaki kod karşılıklarının yazılması yeterli.

Tek istisna: "Ust-Luga – Reshetnikova Office" ile "Ust-Luga - St. Petersburg Office" gibi küçük
yazım farkları var, bunlar eşleme tablosunda ele alınıyor.

---

## 6. Elle yapılan dağıtımda muhtemel hatalar

Otomasyonun asıl değeri sadece zaman kazandırmak değil, elle yapılan işi denetlemek. Temmuz
dosyasında personel verisiyle çelişen kayıtlar buldum:

| Kişi | Elle yazılan şantiye | Personel verisine göre | Sicil |
|---|---|---|---|
| Coskun Emre | RHI | GPP Project, UST LUGA | 512495 |
| Irmak Mehmet Veysi (YRMAK MEKHMET VEISI olarak yazılmış) | RENSERVIS | Regional Management RHI | 300973 |
| Celenligil ailesi | RHI | ALNG2-GBS Project | 423806 |

Bunlar meşru bir iş kuralından da kaynaklanıyor olabilir. Örneğin kişi UST LUGA bordrolu olup
masrafı RHI'ya yazılıyor olabilir. Ama şu an bu kural hiçbir yerde yazılı değil, kişilerin
kafasında. Otomasyon bu farkları her ay listeleyecek ve kural netleştikçe sisteme yazılacak.

---

## 7. En yüksek etkili tek hamle

Şu an tüm zorluk isimle eşleştirmeden geliyor. Bunu bitirecek tek bir değişiklik var:

**Seyahat acentesinden aylık dökümüne TC kimlik veya pasaport numarası kolonu eklemesini isteyin.**

Acente bu bilgiye zaten sahip, çünkü bilet kesmek için pasaport gerekiyor. Tek kolon eklemeleri
eşleştirme problemini büyük ölçüde bitirir. Aynı talep Energo tarafına da yapılabilir.

İkinci hamle: personel ana verisine TC kimlik numarası kolonu eklemek. Şu an sağlık listesinde
ve arabuluculuk dosyasında TC var ama personel verisinde yok, dolayısıyla köprü kurulamıyor.
Bu köprü kurulursa sağlık listesindeki 50 kişi anında sicile bağlanır.

---

## 8. Bu neden AI'a değil masaüstü otomasyonuna ait bir iş

Sorunuzun ikinci kısmı buydu. Cevap: haklısınız, bu iş AI'a ait değil.

Sebepleri:

1. **Problem deterministik.** İsim normalizasyonu, transliterasyon kuralları, tarihe göre dönem
   seçimi, hepsi kural tabanlı. Her ay aynı girdiye aynı çıktıyı vermesi gerekiyor. AI'ın her
   seferinde biraz farklı cevap verme ihtimali burada bir kusur, özellik değil.

2. **Veri kişisel ve hassas.** 24 bin kişinin adı, doğum tarihi, çıkış sebebi. Bunu her ay dışarı
   göndermek gereksiz risk.

3. **Denetlenebilirlik gerekiyor.** Finans ekibi "bu satır neden bu projeye yazıldı" sorusuna
   cevap verebilmeli. Kural tabanlı sistem bunu satır satır açıklayabilir.

4. **Öğrenme AI gerektirmiyor.** Bir kişi bir kez elle eşlenince bu bilgi bir tabloya yazılır ve
   bir daha sorulmaz. Alias tablosu her ay büyür, sistem her ay daha az soru sorar. Bu klasik
   bir referans tablosu, model eğitimi değil.

Sonuç: masaüstünde çalışan, internetsiz, Python tabanlı bir uygulama doğru çözüm.
Uygulamanın kendisi bu deponun içinde, kullanım için README.md dosyasına bakın.

---

## 9. Beklenen çalışma düzeni

Aylık akış şöyle olmalı:

1. Finans ekibi gelen fatura dosyalarını bir klasöre atar.
2. Uygulama açılır, personel verisi yüklenir, dosyalar işlenir.
3. Çıktı üç gruba ayrılır: otomatik eşleşen, incelenmesi gereken, eşleşmeyen.
4. İncelenmesi gereken satırlar arayüzde tek tek onaylanır veya düzeltilir.
5. Her düzeltme alias tablosuna yazılır, bir dahaki ay o kişi otomatik eşleşir.
6. Excel çıktısı muhasebeye gider.

Yeni ay geldiğinde giriş çıkış dosyasına yeni dönemi eklemeniz yeterli. Uygulama dönem
aralığını kendisi tespit ediyor.

İlk aylarda inceleme kuyruğu uzun olacak. Alias tablosu doldukça kısalacak. Kalıcı olarak
incelemede kalacak tek grup aile bireyleri ve dış danışmanlar, çünkü onlar için insan kararı
gerçekten gerekiyor.

---

## 10. Outlook mesajlarini dogrudan okuma

Finans ekibi dosyalari tek tek gondermiyor, Outlook mesaji olarak iletiyor. Ornek dosyada
yapi su sekilde: ana mail icinde iki ekli mail, onlarin icinde zip arsivleri, arsivlerin
icinde Excel dosyalari ve daha fazla mail.

Uygulama bu agaci sonuna kadar yuruyor. Ornek mesaj uzerinde test edildi ve 12 tablo dosyasinin
tamamini cikardi. Turkce ve Kiril karakterli dosya adlari, zip araclarinin urettigi
`#U0131` bicimindeki kacislar ve `>>: Konu` gibi Outlook ek adlari dogru cozuluyor.

Yani kullanici Outlook'tan gelen mesaji dogrudan uygulamaya birakabiliyor. Elle ek acma,
zip cikarma ve klasor duzenleme adimlari ortadan kalkiyor.

Guvenlik notu: arsiv acma sirasinda yol gecisi (path traversal) ve arsiv bombasi kontrolleri
uygulaniyor. Mail imzalarindaki logo gorselleri atlaniyor.
