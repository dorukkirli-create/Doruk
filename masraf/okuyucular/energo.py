"""Energo yansitma dosyalarinin okuyuculari.

Dort ayri sablon, dort ayri fonksiyon:

    assessment_oku()      Assessment/degerlendirme yansitmasi. 'Kisi Listesi'
                          sayfasindaki 'Katilimci' kolonu, kisi basina
                          'Energo Payi' tutari.
    arabulucu_oku()       Arabuluculuk yansitmasi. 'Kisi Listesi' sayfasinda
                          PERSONEL + PERSONEL T.C. + PROJE bulunur; tutar
                          'Fatura Detay' sayfasindan sirket bazinda dagitilir.
    saglik_oku()          Saglik kontrol listesi. ADI SOYADI + TCKN + SANTIYE.
    koc_katilimci_oku()   Koc Universitesi katilimci listesi. ID kolonu
                          dogrudan SICIL NUMARASIDIR (en guvenilir eslesme).

Ortak kurallar: sayfa ADIYLA degil ICERIGIYLE bulunur (baslik satirinda
kisi kolonu + sablona ozgu kolonlar olan sayfa); kolon adlari ASCII katlanmis
'iceren' karsilastirma ile esnek cozulur ('Katilimci Adi' de 'Katilimci'
sayilir), kisi adi bos olan ozet/toplam satirlari atlanir, TCKN 11 haneli
rakam olarak dogrulanir, sicil metne cevrilip '.0' eki atilir.

Olculdu: 'Kisi Listesi' sayfasi 'Liste' olarak yeniden adlandirilinca ya da
'Paket' basligi 'Program' olunca dosya sessizce genel okuyucuya dusuyor, TL
'Toplam' kolonu USD gibi dagitiliyordu (250.260 yerine 6.505,64 USD).
"""

from __future__ import annotations

import re

from pathlib import Path
from typing import Any, Iterable, Sequence

from masraf.kayit import sicil_normalize
from masraf.modeller import GiderSatiri
from masraf.okuyucular.genel import (
    BASLIK_ARAMA_SINIRI,
    baslik_satiri_bul,
    calisma_oku,
    dolu_hucre_sayisi,
    hucre_metni,
    hucre_sayisi,
    hucre_tarihi,
    kolon_anahtari,
    kolon_ara,
    kolon_haritasi,
    sayfa_sec,
    tckn_normalize,
)

__all__ = [
    "assessment_oku",
    "arabulucu_oku",
    "saglik_oku",
    "koc_katilimci_oku",
]

# Sirket kisaltmalarinin kanonik karsiliklari. 'Fatura Detay' sayfasindaki
# 'Masraf yeri' kodlari ile 'Kisi Listesi' sayfasindaki 'Ilgili Sirket'
# degerlerini ayni anahtara indirger.
_SIRKET_ESLERI: dict[str, str] = {
    "rhi": "RHI",
    "rc": "RC",
    "rc peter": "RC",
    "rc moskova": "RC",
    "rsd": "RSD",
    "renstroydetal": "RSD",
    "renservis": "RSS",
    "rs": "RSS",
    "rss": "RSS",
}


def _sirket_anahtari(deger: Any) -> str | None:
    """Sirket kodunu kanonik anahtara cevirir ('renstroydetal' -> 'RSD')."""
    metin = hucre_metni(deger)
    if metin is None:
        return None
    anahtar = kolon_anahtari(metin)
    return _SIRKET_ESLERI.get(anahtar, anahtar.upper() or None)


def _sayfa_satirlari(
    calisma: Any, *aday_adlar: str
) -> tuple[str | None, list[list[Any]]]:
    """Aday adlarla eslesen sayfayi bulur; bulamazsa ilk sayfayi dondurur."""
    ad = sayfa_sec(calisma.sayfa_adlari, *aday_adlar)
    if ad is None:
        ad = calisma.sayfa_adlari[0] if calisma.sayfa_adlari else None
    return ad, (calisma.satirlar(ad) if ad else [])


def _baslik_ara(
    calisma: Any,
    gerekli: Sequence[Sequence[str]],
    secmeli: Sequence[Sequence[str]] = (),
    secmeli_en_az: int = 0,
    haric_sayfalar: Iterable[str] = (),
) -> tuple[str | None, int, dict[str, int]]:
    """Sayfayi ve baslik satirini ADIYLA degil ICERIGIYLE bulur.

    Her sayfanin ilk BASLIK_ARAMA_SINIRI satiri denenir: 'gerekli' gruplarinin
    HER BIRINDEN en az bir kolon adi (normalize edilmis, iceren eslesme) ve
    'secmeli' gruplarindan en az 'secmeli_en_az' tanesi ayni satirda
    bulunuyorsa o satir basliktir. Sayfalar calisma kitabi sirasiyla denenir,
    ilk eslesen kazanir.

    Boylece tedarikci sayfayi 'Kisi Listesi' yerine 'Liste' diye adlandirsa,
    'Katilimci' basligini 'Katilimci Adi' yapsa da sablon taninir; ama kisi
    kolonu olmayan 'Fatura Detay' gibi ozet sayfalari secilmez.

    Returns:
        (sayfa adi, baslik satiri indeksi, kolon haritasi); bulunamazsa
        (None, -1, {}).
    """
    haric = set(haric_sayfalar)
    for sayfa_adi in calisma.sayfa_adlari:
        if sayfa_adi in haric:
            continue
        satirlar = calisma.satirlar(sayfa_adi)
        for i in range(min(len(satirlar), BASLIK_ARAMA_SINIRI)):
            harita = kolon_haritasi(satirlar[i])
            if len(harita) < 2:
                continue
            if not all(kolon_ara(harita, *grup) is not None for grup in gerekli):
                continue
            bulunan = sum(1 for grup in secmeli if kolon_ara(harita, *grup) is not None)
            if bulunan < secmeli_en_az:
                continue
            return sayfa_adi, i, harita
    return None, -1, {}


def _kolon_etiketleri(*gruplar: Sequence[str]) -> str:
    """Kullaniciya gosterilecek 'beklenen kolonlar' metni."""
    return " ve ".join("'" + "' / '".join(g) + "'" for g in gruplar)


def _hucre_alici(satir: list[Any]):
    """Satir icin guvenli indeksli hucre okuyucu uretir."""

    def al(i: int | None) -> Any:
        if i is None or i >= len(satir):
            return None
        return satir[i]

    return al


# --------------------------------------------------------------------------
# 1) Assessment yansitma
# --------------------------------------------------------------------------

_FATURA_NO_DESENI = re.compile(r"\b([A-Z]{2,4}\d{8,})\b")


#: Assessment fatura detay sayfasinin ayirt edici kolonlari.
_ASS_DETAY_PAY = ("energo payi", "pay")
_ASS_DETAY_METIN = ("metin", "aciklama", "fatura")


def _assessment_fatura_ozeti(calisma: Any, kisi_sayfasi: str | None = None) -> dict[str, float]:
    """Fatura detay sayfasindan fatura numarasi bazinda Energo Payi toplamini cikarir.

    Sayfa adiyla degil icerigiyle bulunur: 'Energo Payi' (ya da 'Pay') ve
    'Metin' kolonlarini birlikte tasiyan ilk sayfa. Kisi listesi sayfasi
    disarida tutulur ('Fatura Numarasi' kolonu 'fatura' adayina uyar).

    Satirlar: belge tarihi | yerel tutar | para birimi | USD | 'USD' | Energo Payi | Metin.
    Fatura numarasi 'Metin' kolonundadir ('ASS2026000002187 - AS ...'); toplam
    satirinda metin bos oldugu icin dogal olarak atlanir.
    """
    ad, baslik_i, harita = _baslik_ara(
        calisma, (_ASS_DETAY_PAY, _ASS_DETAY_METIN),
        haric_sayfalar=[kisi_sayfasi] if kisi_sayfasi else (),
    )
    if ad is None:
        return {}
    satirlar = calisma.satirlar(ad)
    i_pay = kolon_ara(harita, *_ASS_DETAY_PAY)
    i_metin = kolon_ara(harita, *_ASS_DETAY_METIN)
    if i_pay is None or i_metin is None:
        return {}
    ozet: dict[str, float] = {}
    for r in range(baslik_i + 1, len(satirlar)):
        al = _hucre_alici(satirlar[r])
        metin = hucre_metni(al(i_metin))
        tutar = hucre_sayisi(al(i_pay))
        if not metin or tutar is None:
            continue
        no = _dosya_adindan_fatura_no(metin) or metin.strip()
        ozet[no] = ozet.get(no, 0.0) + tutar
    return ozet


def _dosya_adindan_fatura_no(dosya_adi: str) -> str | None:
    """'ASS2026000002867 300620261013 ENERGO Fatura Detayi.xlsx' -> 'ASS2026000002867'."""
    m = _FATURA_NO_DESENI.search(dosya_adi or "")
    return m.group(1) if m else None


#: Assessment kisi listesi: kisi kolonu ZORUNLU, sablona ozgu kolonlardan en
#: az ikisi bulunmali. Tek bir ayirt edici kolon (orn. arabuluculuk
#: listesindeki 'Fatura No') bu sablona benzetmeye yetmez.
_ASS_KISI = ("katilimci", "personel", "ad soyad", "adi soyadi", "isim")
_ASS_AYIRT_EDICI: tuple[tuple[str, ...], ...] = (
    ("paket",), ("energo payi", "pay"), ("uygulama turu",), ("uygulama yeri",),
    ("yansitma",), ("firma yetkilisi",), ("fatura numarasi", "fatura no"),
)


def assessment_oku(yol: str | Path) -> list[GiderSatiri]:
    """Assessment yansitma dosyasinin kisi listesi sayfasini okur.

    Sayfa adiyla degil icerigiyle bulunur: baslik satirinda 'Katilimci'
    (ya da 'Personel' / 'Ad Soyad') ve 'Paket' / 'Energo Payi' / 'Uygulama
    Turu' gibi sablona ozgu kolonlardan en az ikisi olan ilk sayfa.

    Kisi adi 'Katilimci' kolonunda AD SOYAD sirasiyla bulunur. Tutar olarak
    kisi satirindaki 'Energo Payi' (USD) alinir; fatura toplami ve USD
    tutari ek sozlugunde saklanir. Masraf merkezi kaynagi 'Yansitma'
    kolonudur (RHI / RSD).
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sayfa_adi, baslik_i, harita = _baslik_ara(
        calisma, (_ASS_KISI,), _ASS_AYIRT_EDICI, secmeli_en_az=2)
    if sayfa_adi is None or baslik_i < 0:
        return []
    satirlar = calisma.satirlar(sayfa_adi)

    i_katilimci = kolon_ara(harita, *_ASS_KISI)
    if i_katilimci is None:
        return []
    i_tarih = kolon_ara(harita, "tarih", icerir=False)
    i_pozisyon = kolon_ara(harita, "pozisyon")
    i_yetkili = kolon_ara(harita, "firma yetkilisi", "yetkili")
    i_tur = kolon_ara(harita, "uygulama turu")
    i_yer = kolon_ara(harita, "uygulama yeri")
    i_paket = kolon_ara(harita, "paket")
    i_fatura_no = kolon_ara(harita, "fatura numarasi", "fatura no")
    i_fatura_tarihi = kolon_ara(harita, "fatura tarihi")
    i_toplam = kolon_ara(harita, "toplam")
    i_usd = kolon_ara(harita, "usd")
    i_pay = kolon_ara(harita, "energo payi", "pay")
    i_yansitma = kolon_ara(harita, "yansitma", "masraf yeri", "sirket")

    # Tedarikci her fatura icin ayrica 'ASS... Fatura Detayi.xlsx' gonderir:
    # ayni katilimcilar, ama TUTAR KOLONU YOK. Tutarlar yansitma dosyasinin
    # 'Kisi Listesi' sayfasindadir. Bu dosya fatura degil, faturanin kisi
    # listesidir; gider satiri uretirse mutabakat 'tutar okunamadi' diye
    # acik kalir. Kutuk olarak isaretlenir, mahsuplasmada yansitma
    # satirlariyla capraz kontrol edilir (bkz. mahsuplasma.detay_kontrolu).
    detay_listesi = i_pay is None and i_usd is None and i_toplam is None
    kaynak_tip = "energo_assessment_detay" if detay_listesi else "energo_assessment"
    dosya_fatura_no = _dosya_adindan_fatura_no(p.name)
    # Yansitma dosyasinin 'Fatura Detay' sayfasi fatura basina Energo Payi'ni
    # beyan eder; kisi satirlarinin toplamiyla kurusuna kadar karsilastirilir.
    fatura_ozeti = _assessment_fatura_ozeti(calisma, sayfa_adi) if not detay_listesi else {}

    sonuclar: list[GiderSatiri] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        al = _hucre_alici(satir)
        katilimci = hucre_metni(al(i_katilimci))
        if katilimci is None:
            continue  # toplam satiri / pivot artigi

        tutar = hucre_sayisi(al(i_pay))
        if tutar is None:
            tutar = hucre_sayisi(al(i_usd))
        fatura_no = hucre_metni(al(i_fatura_no)) or dosya_fatura_no

        fatura_tarihi = hucre_tarihi(al(i_fatura_tarihi), calisma.datemode)
        katilim_tarihi = hucre_tarihi(al(i_tarih), calisma.datemode)

        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip=kaynak_tip,
                satir_no=r + 1,
                belge_tarihi=fatura_tarihi or katilim_tarihi,
                aciklama=" | ".join(
                    m
                    for m in (
                        katilimci,
                        hucre_metni(al(i_paket)),
                        hucre_metni(al(i_tur)),
                        fatura_no,
                    )
                    if m
                ),
                kisi_ham=katilimci,
                sicil_ham=None,
                tckn_ham=None,
                tutar=tutar,
                para_birimi="USD",
                masraf_merkezi_kaynak=hucre_metni(al(i_yansitma)),
                gider_tipi="Egitim",
                ek={
                    "sayfa": sayfa_adi,
                    "katilim_tarihi": katilim_tarihi,
                    "pozisyon": hucre_metni(al(i_pozisyon)),
                    "firma_yetkilisi": hucre_metni(al(i_yetkili)),
                    "uygulama_turu": hucre_metni(al(i_tur)),
                    "uygulama_yeri": hucre_metni(al(i_yer)),
                    "paket": hucre_metni(al(i_paket)),
                    "fatura_no": fatura_no,
                    "fatura_toplam": hucre_sayisi(al(i_toplam)),
                    "fatura_usd": hucre_sayisi(al(i_usd)),
                    "tutar_yontemi": (
                        "fatura detay listesi; tutar yansitma dosyasinda"
                        if detay_listesi else "kisi satirindaki Energo Payi"
                    ),
                    **({"fatura_ozeti": dict(fatura_ozeti),
                        "beyan_yontemi": "'Fatura Detay' sayfasindaki fatura bazli Energo Payi"}
                       if fatura_ozeti else {}),
                },
            )
        )
    return sonuclar


# --------------------------------------------------------------------------
# 2) Arabuluculuk yansitma
# --------------------------------------------------------------------------

#: Arabuluculuk fatura detay sayfasinin ayirt edici kolonlari. Pay kolonu iki
#: kademede aranir: once 'Energo Payi' / 'Pay (USD)', o yoksa 'Fatura Tutari
#: (USD)'. Tek listede aranamaz: kolon_ara once BUTUN adaylarin tam
#: eslesmesine bakar, 'Fatura Tutari ( USD )' tam eslesip 'Pay (USD)'nin
#: onune gecer ve fatura tutari pay diye dagitilir (olculdu: 15 yerine 5,25).
_ARA_DETAY_YER = ("masraf yeri", "masraf merkezi", "sirket")
_ARA_DETAY_PAY = ("energo payi", "pay")
_ARA_DETAY_PAY_YEDEK = ("fatura tutari usd",)


def _arabulucu_fatura_ozeti(
    calisma: Any, kisi_sayfasi: str | None = None
) -> tuple[dict[str, float], str | None]:
    """Fatura detay sayfasindan sirket bazinda Energo Payi toplamini cikarir.

    Sayfa adiyla degil icerigiyle bulunur: 'Masraf yeri' (ya da 'Sirket')
    ve 'Energo Payi' (ya da 'Pay') kolonlarini birlikte tasiyan ilk sayfa;
    kisi listesi sayfasi disarida tutulur. Ozet/pivot satirlari 'Masraf
    yeri' bos oldugu icin dogal olarak atlanir.

    Returns:
        (sirket -> toplam, sebep). Sayfa bulunamadiysa ya da hicbir satirda
        tutar okunamadiysa sebep, finansciya ne aranip ne bulunamadigini
        soyler; satir aciklamasina bu metin yazilir. Olculdu: 'Fatura Detay'
        sayfasi 'Ozet' olunca 25 kisinin tutari None kaliyor ve aciklama
        yaniltici bicimde 'eslesen masraf yeri yok' diyordu.
    """
    ad, baslik_i, harita = _baslik_ara(
        calisma, (_ARA_DETAY_YER, _ARA_DETAY_PAY + _ARA_DETAY_PAY_YEDEK),
        haric_sayfalar=[kisi_sayfasi] if kisi_sayfasi else (),
    )
    if ad is None:
        digerleri = [a for a in calisma.sayfa_adlari if a != kisi_sayfasi]
        return {}, (
            "fatura detay sayfasi bulunamadi: beklenen kolonlar "
            + _kolon_etiketleri(("Masraf yeri", "Sirket"), ("Energo Payi", "Pay"))
            + " birlikte hicbir sayfada yok"
            + (f" (bakilan sayfalar: {', '.join(digerleri)})" if digerleri else "")
        )
    satirlar = calisma.satirlar(ad)
    i_yer = kolon_ara(harita, *_ARA_DETAY_YER)
    i_pay = kolon_ara(harita, *_ARA_DETAY_PAY)
    if i_pay is None:
        i_pay = kolon_ara(harita, *_ARA_DETAY_PAY_YEDEK)
    if i_yer is None or i_pay is None:  # _baslik_ara garanti eder; savunma
        return {}, f"fatura detay sayfasi ('{ad}') bulundu ama kolonlari cozulemedi"

    ozet: dict[str, float] = {}
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        al = _hucre_alici(satir)
        anahtar = _sirket_anahtari(al(i_yer))
        tutar = hucre_sayisi(al(i_pay))
        if anahtar is None or tutar is None:
            continue
        ozet[anahtar] = ozet.get(anahtar, 0.0) + tutar
    if not ozet:
        return {}, (
            f"fatura detay sayfasi ('{ad}') bulundu ama hicbir satirda masraf yeri "
            "ve tutar birlikte okunamadi (hucreler bos ya da formul sonucu "
            "kaydedilmemis olabilir)"
        )
    return ozet, None


def _kurusa_bol(toplam: float, adet: int) -> list[float]:
    """toplam'i adet kisiye kurus hassasiyetinde, artiksiz boler.

    Sonuc listesinin toplami round(toplam, 2)'ye ESITTIR. Artik kuruslar
    listenin basindaki kisilere birer birer eklenir (en buyuk kalan yontemi
    esit paylarda buna indirgenir).
    """
    kurus = int(round(toplam * 100))
    taban, artik = divmod(kurus, adet)
    return [(taban + (1 if i < artik else 0)) / 100 for i in range(adet)]


#: Arabuluculuk kisi listesi: kisi kolonu ve (arabulucu | TCKN | proje)
#: zorunlu; sablona ozgu kolonlardan en az ikisi bulunmali. 'Ad Soyad' +
#: 'Arabulucu Ucreti' basligi tasiyan sade bir gider dosyasi bu sablona
#: benzetilmez (tek ayirt edici kolon), genel okuyucuya birakilir.
_ARA_KISI = ("personel", "ad soyad", "adi soyadi", "katilimci", "isim")
_ARA_ZORUNLU_BIRI = ("arabulucu", "personel t c", "tckn", "tc kimlik no", "proje")
_ARA_AYIRT_EDICI: tuple[tuple[str, ...], ...] = (
    ("arabulucu",), ("personel t c", "tckn", "tc kimlik no"),
    ("proje", "santiye", "masraf yeri"), ("ilgili sirket", "sirket"),
    ("yetkili",), ("fatura no", "fatura numarasi"),
)


def arabulucu_oku(yol: str | Path) -> list[GiderSatiri]:
    """Arabuluculuk yansitma dosyasinin kisi listesi sayfasini okur.

    Bu sablonda PERSONEL T.C. (TCKN) ve PROJE kolonlari vardir; PROJE
    degerleri personel ana verisindeki 'Gorev Yeri' degerleriyle birebir
    ayni DEGILDIR, esleme tablosu gerektirir (veri/masraf_merkezi_haritasi.csv).

    Kisi listesi sayfasi adiyla degil icerigiyle bulunur: baslik satirinda
    kisi kolonu ('Personel' / 'Ad Soyad'), 'Arabulucu' / 'Personel T.C.' /
    'Proje' kolonlarindan biri ve sablona ozgu kolonlardan en az ikisi.

    Kisi listesi sayfasinda tutar kolonu bulunmadigi icin tutar, fatura
    detay sayfasindaki sirket bazli Energo Payi toplaminin o sirkete ait
    kisi sayisina esit bolunmesiyle hesaplanir; yontem ek['tutar_yontemi']
    icinde acikca belirtilir. Fatura detay sayfasi bulunamazsa ya da
    eslesen masraf yeri yoksa tutar None kalir ve sebep ayni alana yazilir.
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sayfa_adi, baslik_i, harita = _baslik_ara(
        calisma, (_ARA_KISI, _ARA_ZORUNLU_BIRI), _ARA_AYIRT_EDICI, secmeli_en_az=2)
    if sayfa_adi is None or baslik_i < 0:
        return []
    satirlar = calisma.satirlar(sayfa_adi)

    i_tckn = kolon_ara(harita, "personel t c", "tckn", "tc kimlik no", "kimlik no")
    i_personel = kolon_ara(harita, *_ARA_KISI)
    if i_personel is not None and i_personel == i_tckn:
        # 'personel' anahtari 'personel t c' kolonuna dusmus olabilir
        i_personel = harita.get("personel")
    if i_personel is None:
        return []
    i_tarih = kolon_ara(harita, "tarih", icerir=False)
    i_yetkili = kolon_ara(harita, "yetkili")
    i_proje = kolon_ara(harita, "proje", "santiye", "masraf yeri")
    i_sirket = kolon_ara(harita, "ilgili sirket", "sirket")
    i_arabulucu = kolon_ara(harita, "arabulucu")
    i_fatura_no = kolon_ara(harita, "fatura no", "fatura numarasi")
    i_fatura_tarihi = kolon_ara(harita, "fatura tarihi")

    # Once kisileri topla, sonra sirket basina esit paylastir.
    ham_satirlar: list[tuple[int, list[Any]]] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        if dolu_hucre_sayisi(satir) == 0:
            continue
        al = _hucre_alici(satir)
        if hucre_metni(al(i_personel)) is None:
            continue
        ham_satirlar.append((r, satir))

    fatura_ozeti, detay_sebebi = _arabulucu_fatura_ozeti(calisma, sayfa_adi)
    sayimlar: dict[str, int] = {}
    for _, satir in ham_satirlar:
        anahtar = _sirket_anahtari(_hucre_alici(satir)(i_sirket))
        if anahtar:
            sayimlar[anahtar] = sayimlar.get(anahtar, 0) + 1

    # Etiket uyusmazligi koprusu. Tedarikci ozet sayfasinda 'RSD' yazar,
    # kisi listesinde 'Renservis' der; ikisi ayni paydir ama sozlukte ayni
    # anahtara dusmez. Sozlugu genisletmek kirilgan: her ay yeni bir yazim
    # gelebilir. Saglam kural: kisisiz kalan TEK bir kova ve kovasiz kalan
    # kisiler varsa, o kisiler o kovanindir. Olculdu: Temmuz 2026'da bu
    # kopru olmadan 76,78 USD sessizce dagilim disinda kaliyordu.
    kisisiz_kovalar = [k for k in fatura_ozeti if not sayimlar.get(k)]
    kovasiz_kisiler = [k for k in sayimlar if k not in fatura_ozeti]
    kopru: dict[str, str] = {}
    if len(kisisiz_kovalar) == 1 and kovasiz_kisiler:
        for k in kovasiz_kisiler:
            kopru[k] = kisisiz_kovalar[0]
        for k in kovasiz_kisiler:
            sayimlar[kisisiz_kovalar[0]] = sayimlar.get(kisisiz_kovalar[0], 0) + sayimlar.pop(k)

    # Esit bolme kurusta kayip yaratir: 1709,91 / 22 = 77,7233 -> 77,72 x 22 =
    # 1709,84; yedi kurus yok olur ve fatura beyanla kapanmaz. Kurus artigi
    # ilk kisilere birer kurus olarak eklenir; boylece kisi tutarlari toplami
    # sirket toplaminin kendisine esittir. Olculdu: Temmuz 2026'da 0,07 USD.
    paylar: dict[str, list[float]] = {
        k: _kurusa_bol(fatura_ozeti[k], n)
        for k, n in sayimlar.items() if k in fatura_ozeti and n
    }
    pay_sirasi: dict[str, int] = {k: 0 for k in paylar}

    sonuclar: list[GiderSatiri] = []
    for r, satir in ham_satirlar:
        al = _hucre_alici(satir)
        personel = hucre_metni(al(i_personel))
        sirket_ham = _sirket_anahtari(al(i_sirket))
        sirket = kopru.get(sirket_ham, sirket_ham)

        tutar: float | None = None
        if detay_sebebi:
            yontem = detay_sebebi
        else:
            yontem = (
                f"fatura detayinda eslesen masraf yeri yok: kisi listesinde "
                f"'{sirket_ham or '-'}' yaziyor, detaydaki masraf yerleri: "
                + ", ".join(sorted(fatura_ozeti))
            )
        if sirket in paylar:
            tutar = paylar[sirket][pay_sirasi[sirket]]
            pay_sirasi[sirket] += 1
            yontem = (
                f"'{sirket}' masraf yeri toplami "
                f"({fatura_ozeti[sirket]:.2f}) / {sayimlar[sirket]} kisi"
            )
            if abs(tutar * sayimlar[sirket] - fatura_ozeti[sirket]) >= 0.005:
                yontem += "; kurus artigi ilk kisilere dagitildi"
            if sirket_ham != sirket:
                yontem += (f"; kisi listesinde '{sirket_ham}' yaziyor, fatura detayinda "
                           f"karsiligi olmayan tek kova '{sirket}' oldugu icin ona baglandi")

        olay_tarihi = hucre_tarihi(al(i_tarih), calisma.datemode)
        fatura_tarihi = hucre_tarihi(al(i_fatura_tarihi), calisma.datemode)

        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="energo_arabulucu",
                satir_no=r + 1,
                belge_tarihi=fatura_tarihi or olay_tarihi,
                aciklama=" | ".join(
                    m
                    for m in (
                        personel,
                        hucre_metni(al(i_proje)),
                        hucre_metni(al(i_arabulucu)),
                        hucre_metni(al(i_fatura_no)),
                    )
                    if m
                ),
                kisi_ham=personel,
                sicil_ham=None,
                tckn_ham=tckn_normalize(al(i_tckn)),
                tutar=tutar,
                para_birimi="USD",
                masraf_merkezi_kaynak=hucre_metni(al(i_proje)),
                gider_tipi="Arabuluculuk",
                ek={
                    "sayfa": sayfa_adi,
                    "olay_tarihi": olay_tarihi,
                    "yetkili": hucre_metni(al(i_yetkili)),
                    "ilgili_sirket": hucre_metni(al(i_sirket)),
                    "sirket_kodu": sirket,
                    "arabulucu": hucre_metni(al(i_arabulucu)),
                    "fatura_no": hucre_metni(al(i_fatura_no)),
                    "tutar_yontemi": yontem,
                    "fatura_ozeti": dict(fatura_ozeti),
                },
            )
        )
    return sonuclar


# --------------------------------------------------------------------------
# 3) Saglik kontrol listesi
# --------------------------------------------------------------------------

def saglik_oku(yol: str | Path, bordrosuz_dahil: bool = False) -> list[GiderSatiri]:
    """Saglik kontrol listesini okur ('BORDROLU LISTE' sayfasi).

    ADI SOYADI (AD SOYAD sirasiyla, Turkce karakterli), TCKN, DOGUM TARIHI ve
    SANTIYE kolonlari bulunur. TCKN sayesinde bu dosya, personel ana
    verisinde bulunmayan kisiler icin EK KISI DEFTERI kaynagidir.

    Args:
        bordrosuz_dahil: True ise 'BORDROSUZ LISTE' sayfasi da okunur.
            Varsayilan False - bordrosuz taseronlar masraf mahsuplastirma
            akisinda ayri ele alinir.
    """
    p = Path(yol)
    calisma = calisma_oku(p)

    hedef_sayfalar: list[str] = []
    bordrolu = sayfa_sec(calisma.sayfa_adlari, "BORDROLU LISTE", "BORDROLU")
    if bordrolu:
        hedef_sayfalar.append(bordrolu)
    if bordrosuz_dahil:
        bordrosuz = sayfa_sec(calisma.sayfa_adlari, "BORDROSUZ LISTE", "BORDROSUZ")
        if bordrosuz and bordrosuz not in hedef_sayfalar:
            hedef_sayfalar.append(bordrosuz)
    if not hedef_sayfalar and calisma.sayfa_adlari:
        hedef_sayfalar.append(calisma.sayfa_adlari[0])

    sonuclar: list[GiderSatiri] = []
    for sayfa_adi in hedef_sayfalar:
        satirlar = calisma.satirlar(sayfa_adi)
        baslik_i = baslik_satiri_bul(satirlar, aranan=("ADI SOYADI", "TCKN", "S.NO"))
        if baslik_i < 0:
            continue
        harita = kolon_haritasi(satirlar[baslik_i])

        i_sno = kolon_ara(harita, "s no", "sno", "sira no")
        i_isim = kolon_ara(harita, "adi soyadi", "ad soyad", "isim")
        if i_isim is None:
            continue
        i_tckn = kolon_ara(harita, "tckn", "tc kimlik no", "tc kimlik", "kimlik no")
        i_dogum = kolon_ara(harita, "dogum tarihi")
        i_ulke = kolon_ara(harita, "ulke")
        i_gorev = kolon_ara(harita, "gorevi", "gorev")
        i_santiye = kolon_ara(harita, "santiye", "proje")
        i_firma = kolon_ara(harita, "firma ekip formen", "firma")
        i_iletisim = kolon_ara(harita, "iletisim bilgileri 1", "iletisim")
        i_talep = kolon_ara(harita, "talep tarihi")
        i_kontrol = kolon_ara(
            harita, "saglik kontrol tarihi", "kontrol tarihi dr", "kontrol tarihi"
        )
        bordrolu_mu = "bordrosuz" not in kolon_anahtari(sayfa_adi)

        for r in range(baslik_i + 1, len(satirlar)):
            satir = satirlar[r]
            al = _hucre_alici(satir)
            isim = hucre_metni(al(i_isim))
            if isim is None:
                continue

            kontrol_tarihi = hucre_tarihi(al(i_kontrol), calisma.datemode)
            talep_tarihi = hucre_tarihi(al(i_talep), calisma.datemode)

            sonuclar.append(
                GiderSatiri(
                    kaynak_dosya=p.name,
                    kaynak_tip="energo_saglik",
                    satir_no=r + 1,
                    belge_tarihi=kontrol_tarihi or talep_tarihi,
                    aciklama=" | ".join(
                        m
                        for m in (
                            isim,
                            hucre_metni(al(i_gorev)),
                            hucre_metni(al(i_santiye)),
                        )
                        if m
                    ),
                    kisi_ham=isim,
                    sicil_ham=None,
                    tckn_ham=tckn_normalize(al(i_tckn)),
                    tutar=None,
                    para_birimi=None,
                    masraf_merkezi_kaynak=hucre_metni(al(i_santiye)),
                    gider_tipi="Saglik",
                    ek={
                        "sayfa": sayfa_adi,
                        "bordrolu": bordrolu_mu,
                        "s_no": hucre_metni(al(i_sno)),
                        "dogum_tarihi": hucre_tarihi(al(i_dogum), calisma.datemode),
                        "ulke": hucre_metni(al(i_ulke)),
                        "gorevi": hucre_metni(al(i_gorev)),
                        "firma": hucre_metni(al(i_firma)),
                        "iletisim": hucre_metni(al(i_iletisim)),
                        "talep_tarihi": talep_tarihi,
                        "kontrol_tarihi": kontrol_tarihi,
                    },
                )
            )
    return sonuclar


# --------------------------------------------------------------------------
# 4) Koc Universitesi katilimci listesi
# --------------------------------------------------------------------------

def koc_katilimci_oku(yol: str | Path) -> list[GiderSatiri]:
    """Koc Universitesi egitim katilimci listesini okur.

    'ID' kolonu dogrudan SICIL NUMARASIDIR; bu dosya en guvenilir eslesmeyi
    saglar. 'Ad Soyad' degerleri SOYAD AD sirasindadir ('Birladean Alexandr').

    Ayni kisi birden fazla katilim tarihi icin tekrar eder; her satir AYRI
    bir GiderSatiri olarak dondurulur (tekillestirme yapilmaz).
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sayfa_adi, satirlar = _sayfa_satirlari(
        calisma, "Sheet1", "Katilimci Listesi", "Kisi Listesi"
    )
    if not satirlar:
        return []

    baslik_i = baslik_satiri_bul(satirlar, aranan=("Ad Soyad", "Alt Fonksiyon", "ID"))
    if baslik_i < 0:
        return []
    harita = kolon_haritasi(satirlar[baslik_i])

    i_id = kolon_ara(harita, "id", "sicil", "sicil no", "personel no")
    i_isim = kolon_ara(harita, "ad soyad", "adi soyadi", "katilimci", "isim")
    if i_isim is None:
        return []
    i_pozisyon = kolon_ara(harita, "pozisyon", "gorev")
    i_fonksiyon = kolon_ara(harita, "alt fonksiyon", "fonksiyon")
    i_tarih = kolon_ara(harita, "katilim tarihi", "tarih")

    sonuclar: list[GiderSatiri] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        al = _hucre_alici(satir)
        isim = hucre_metni(al(i_isim))
        sicil = sicil_normalize(al(i_id)) or None
        if isim is None and sicil is None:
            continue

        katilim = hucre_tarihi(al(i_tarih), calisma.datemode)
        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="koc_katilimci",
                satir_no=r + 1,
                belge_tarihi=katilim,
                aciklama=" | ".join(
                    m
                    for m in (
                        isim,
                        hucre_metni(al(i_pozisyon)),
                        hucre_metni(al(i_fonksiyon)),
                    )
                    if m
                ),
                kisi_ham=isim,
                sicil_ham=sicil,
                tckn_ham=None,
                tutar=None,
                para_birimi=None,
                masraf_merkezi_kaynak=None,
                gider_tipi="Egitim",
                ek={
                    "sayfa": sayfa_adi,
                    "pozisyon": hucre_metni(al(i_pozisyon)),
                    "alt_fonksiyon": hucre_metni(al(i_fonksiyon)),
                    "katilim_tarihi": katilim,
                },
            )
        )
    return sonuclar
