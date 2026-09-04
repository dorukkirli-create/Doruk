# veri/ dizini

Bu dizin uygulamanin ogrenen defterlerini tutar. Iceriginin cogu
**kisisel veri** oldugu icin repoya gonderilmez, `.gitignore` disarida
birakir. Dosyalar yoksa uygulama ilk calismada baslik satirlariyla
olusturur.

| Dosya | Icerik | Repoda mi |
|---|---|---|
| `masraf_merkezi_haritasi.csv` | Gorev yeri -> finans masraf merkezi kodu | evet, yapilandirma |
| `aliases.csv` | Elle onaylanmis isim -> sicil eslesmeleri | hayir |
| `ek_kisiler.csv` | Yardimci listelerden gelen kisiler, TC kimlik icerir | hayir |
| `tckn_sicil.csv` | TC kimlik -> sicil koprusu | hayir |
| `harici_kisiler.csv` | Calisan olmayan kisiler, dis danismanlar | hayir |

## masraf_merkezi_haritasi.csv nasil doldurulur

`masraf_merkezi_kodu` kolonundaki degerler simdilik gorev yerinin
kisaltmasidir. Finans ekibinin kendi kod setiyle degistirilmelidir.
Uygulama haritada olmayan bir gorev yeri gordugunde satiri uyari ile
isaretler ve gorev yerini oldugu gibi kullanir.

## Defterler nasil buyur

Inceleme ekraninda bir satiri bir sicile bagladiginizda kayit
`aliases.csv` dosyasina yazilir. Ayni isim bir daha geldiginde otomatik
eslesir. Ayni sekilde "calisan degil" dediginiz kisiler
`harici_kisiler.csv` dosyasina gider.

Bu dosyalari yedeklemek isterseniz sirket ici bir paylasima kopyalayin,
genel bir depoya koymayin.
