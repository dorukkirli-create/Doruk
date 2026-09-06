"""Antik / Yuzyil Travel seyahat acentesi dosyalarinin okuyuculari.

Iki farkli dosya ailesi:

    antik_cari_oku()        ham cari hareket dokumu (.xls, xlrd ile okunur).
                            Masraf merkezi kolonu YOKTUR; kisi adi 'Aciklama'
                            metninin icine gomuludur.

    yuzyil_dagitilmis_oku() ayni ayin elle santiyeye dagitilmis hali (.xlsx).
                            'SANTIYESI' kolonu doldurulmustur; otomasyonun
                            dogrulugunu olcmek icin DOGRULUK REFERANSI olarak
                            kullanilir, ana is akisinda girdi degildir.

Kisi adi cikarma islemi gider tipine gore ayri fonksiyonlara bolunmustur;
her biri tek basina test edilebilir (_bilet_kisi, _otel_kisi, _vize_kisi,
_bagaj_kisi, _diger_kisi).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from masraf.metin import OTEL_ANAHTARLARI, ascii_katla, kisi_metnini_temizle, tr_buyuk
from masraf.modeller import GiderSatiri
from masraf.okuyucular.genel import (
    Calisma,
    baslik_satiri_bul,
    calisma_oku,
    dolu_hucre_sayisi,
    hucre_metni,
    hucre_sayisi,
    hucre_tarihi,
    kolon_ara,
    kolon_haritasi,
)

__all__ = ["antik_cari_oku", "yuzyil_dagitilmis_oku"]


# --------------------------------------------------------------------------
# Kisi adi olamayacak kelimeler
# --------------------------------------------------------------------------

# Bir isim tokeni bu kumedeyse orasi isim degil hizmet aciklamasidir; token
# listesi o noktadan itibaren kesilir. Kume 'kuyruk kirpma' icin kullanilir,
# bu yuzden ILK gecisin konumu onemlidir.
_HIZMET_KELIMELERI: frozenset[str] = frozenset({
    "BILET", "BEDELI", "BEDEL", "UCRETI", "UCRET", "TUTARI",
    "OTEL", "HOTEL", "KONAKLAMA", "GECE", "GECELIK", "ODA",
    "VIZE", "VIZESI", "EVIZE", "KONSOLOSLUK", "ELCILIK", "BASVURU",
    "RUSYA", "FEDERASYONU", "TURISTIK", "DAVETIYE", "DAVET", "MEKTUBU",
    "TRANSFER", "HAVALIMANI", "HAVAALANI", "TERMINAL", "MERKEZ",
    "OTOBUS", "ARAC", "KIRALAMA", "KIRA", "SOFOR",
    "CENAZE", "CELENK", "CELENGI", "GONDERIMI", "GONDERIM",
    "TOPLANTI", "ORGANIZASYON", "ORGANIZASYONU", "PAKET", "PAKETI",
    "CATERING", "YEMEK", "SALON",
    "DIGER", "GIDER", "GIDERLER", "HIZMET", "HIZMETLER", "SERVIS",
    "YURT", "ICI", "DISI", "YURTICI", "YURTDISI",
    "EKSTRA", "BAGAJ", "TASINDI", "TARAFINDAN",
    "REZERVASYON", "DEGISIKLIK", "IPTAL", "IADE", "KOMISYON", "FARK",
    "SIGORTA", "ASISTAN", "ASISTANS",
    "FATURA", "EVRAK", "ADET", "TOPLAM", "SARJ", "EDILECEK",
})

# Hem KISI ADI hem hizmet kelimesi olabilenler. 'IKRAM' Turkiye'de gercek bir
# addir (olculdu: personel verisinde 7 kisi); kosulsuz hizmet kelimesi
# sayilinca bu kisilerin adi tek tokene dusuyor ve ESLESMEDI cikiyordu.
# Kural: yalnizca en az iki ad tokeninden SONRA ve tipik hizmet baglaminda
# (sonraki token asagidaki kumedeyse: 'IKRAM BEDELI', 'IKRAM HIZMETI')
# kesilir; tek basina ya da adin icinde gecince ad olarak kalir.
_BAGLAMLI_HIZMET_KELIMELERI: dict[str, frozenset[str]] = {
    "IKRAM": frozenset({
        "BEDELI", "BEDEL", "UCRETI", "UCRET", "TUTARI", "HIZMETI", "HIZMET",
        "HIZMETLERI", "GIDERI", "GIDERLERI", "SERVISI", "ORGANIZASYONU",
        "PAKETI", "CATERING",
    }),
}

# Bilet numarasi oneki: iki harfli havayolu kodu + uzun rakam dizisi.
_RE_BILET_ONEK = re.compile(r"^\s*([A-Z]{2})\s*(\d{6,})\s+(?P<kalan>.+)$")
# Cinsiyet/yolcu tipi isareti: \M \F \I \C  (ters bolu + tek harf)
_RE_YOLCU_ISARETI = re.compile(r"\\\s*[A-Z]\b")
# Unvan tokeni
_RE_UNVAN = re.compile(r"\s+(?:MR|MRS|MS|MSTR|MISS|CHD|INF)\b")
# IATA guzergah zinciri: IST-CDG, KYA-SAW-LED
_RE_IATA_ZINCIR = re.compile(r"\b[A-Z]{3}(?:\s*-\s*[A-Z]{3})+\b")
# Tarih: 11.07.2026 veya 11/07/2026
_RE_TARIH = re.compile(r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b")
# Otel adinin basladigi anahtar kelimeler
_RE_OTEL_ANAHTAR = re.compile(
    r"\b(?:" + "|".join(re.escape(k) for k in OTEL_ANAHTARLARI) + r")\b"
)
# Coklu kisi ayiraclari (otel satirlarinda)
_RE_KISI_AYIRAC = re.compile(r"[;&+]|\s/\s")
# Bagaj kalibi. Kaynak veride 'GUNDOGDUTARAFINDAN' gibi bosluk dusmus
# yazimlar goruldugu icin ayirac '\s*' (sifir veya daha fazla bosluk).
_RE_BAGAJ = re.compile(
    r"EKSTRA\s+BAGAJ(?:\s+UCRETI)?\s+(?P<isim>.+?)\s*TARAFINDAN", re.S
)

# Isim tokenine yapisik gelebilen uzun hizmet kelimeleri. Sadece yeterince
# uzun ve tek anlamli kelimeler burada olmalidir; kisa ekler ('ICI', 'NO')
# gercek soyadlarini bolebilecegi icin KASITLI olarak disaridadir.
_YAPISIK_KUYRUKLAR: tuple[str, ...] = (
    "TARAFINDAN", "KONAKLAMA", "TASINDI", "BEDELI", "UCRETI", "FEDERASYONU",
)
# Dokumun kendi toplam/ozet satirinin ETIKETI. Hucrenin TAMAMI bu kaliba
# uymali: 'TOPLAM', 'GENEL TOPLAM', 'BAKIYE', 'TOPLAM USD', 'TOPLAM BORC:'.
# Bir veri satirinin aciklamasinda gecen '(TOPLAM 2 KISI)' uymaz.
_RE_TOPLAM_ETIKETI = re.compile(
    r"^\s*(?:GENEL\s+|ARA\s+|NET\s+)?(?:TOPLAM|TOTAL|BAKIYE)"
    r"(?:\s+(?:BORC|ALACAK|BAKIYE|TUTAR|TUTARI|USD|TL|TRY|EUR|RUB|GENEL|NET))*"
    r"\s*:?\s*$"
)
# 'RHI 1/3- RENSTROYDETAL 2/3' bicimli paylasim. Tarih ('GIRIS 10/07/2026') ve
# saat ('10/07 14:30') kesirleri disarida tutulur: kesirden once rakam ya da
# '/' bulunamaz, kesirden sonra '/rakam', ':rakam' ya da '.rakam' gelemez.
# Eski desen 'GIRIS 10/07/2026' notundan pay 10 / bolen 7 = 1,43 uretiyordu.
_RE_PAYLASIM = re.compile(
    r"(?<![\d/])([A-Z][A-Z0-9]*(?:\s+[A-Z0-9]+)*?)\s*(\d{1,2})\s*/\s*(\d{1,2})(?!\s*[/:.]\s*\d)(?!\d)"
)
#: Paylasim boleni en fazla bu olabilir (1/12 en ince paylasim). Pay bolenden
#: buyuk olamaz; '10/07' gibi tarih parcalari ve '2/25' gibi sayimlar
#: paylasim degildir.
_PAYLASIM_AZAMI_BOLEN = 12

# 'Islem' kolonundan gider tipine esleme (normalize anahtar -> tip).
_ISLEM_TIPLERI: dict[str, str] = {
    "bilet islem": "Bilet",
    "bilet islemleri": "Bilet",
    "otel islemleri": "Otel",
    "otel islem": "Otel",
    "vize islemleri": "Vize",
    "vize islem": "Vize",
    "diger hizmetler": "Diger",
    "transfer": "Diger",
}


def _fold(metin: Any) -> str:
    """Metni buyuk harf + ASCII bicime katlar (Turkce guvenli)."""
    if metin is None:
        return ""
    return ascii_katla(tr_buyuk(str(metin)))


def _yapisik_kuyruk_ac(token: str) -> str:
    """Isim tokenine yapisik yazilmis hizmet kelimesini ayirir.

    Kaynak veride bosluk dusmus yazimlar vardir:
    'GUNDOGDUTARAFINDAN' -> 'GUNDOGDU'. Geriye en az 3 harf kalmiyorsa
    token oldugu gibi birakilir.
    """
    for kuyruk in _YAPISIK_KUYRUKLAR:
        if token.endswith(kuyruk) and len(token) - len(kuyruk) >= 3:
            return token[: -len(kuyruk)]
    return token


def _kuyruk_kirp(isim: str) -> str:
    """Isim tokenlarini ilk hizmet kelimesinde keser.

    >>> _kuyruk_kirp("OMER CAN CETIR KONSOLOSLUK UCRETI")
    'OMER CAN CETIR'
    >>> _kuyruk_kirp("ALI GUNDOGDUTARAFINDAN")
    'ALI GUNDOGDU'
    >>> _kuyruk_kirp("ALI VELI IKRAM BEDELI")
    'ALI VELI'
    >>> _kuyruk_kirp("IKRAM YILMAZ")
    'IKRAM YILMAZ'
    """
    tokenlar: list[str] = []
    parcalar = isim.split()
    for k, token in enumerate(parcalar):
        if token in _HIZMET_KELIMELERI:
            break
        baglam = _BAGLAMLI_HIZMET_KELIMELERI.get(token)
        if (baglam is not None and len(tokenlar) >= 2
                and k + 1 < len(parcalar) and parcalar[k + 1] in baglam):
            break  # 'AD SOYAD IKRAM BEDELI': ad burada biter
        acik = _yapisik_kuyruk_ac(token)
        tokenlar.append(acik)
        if acik != token:
            break  # yapisik hizmet kelimesi bulundu, isim burada biter
    return " ".join(tokenlar)


def _toplam_satiri_mi(
    satir: list[Any], al: Any, i_tarih: int | None, i_evrak: int | None,
    i_aciklama: int | None,
) -> bool:
    """Satir dokumun kendi TOPLAM / BAKIYE satiri mi?

    Iki sart birden aranir:
      1. Kimlik hucreleri bos: islem tarihi ve evrak no yok; aciklama ya bos
         ya da kendisi toplam etiketi ('TOPLAM').
      2. Satirdaki bir hucrenin TAMAMI toplam etiketi (bkz. _RE_TOPLAM_ETIKETI).

    Onceki kural satirin HERHANGI bir hucresinde 'TOPLAM' kelimesini
    ariyordu; 'AHMET ... (TOPLAM 2 KISI)' gibi aciklama tasiyan bir veri
    satiri toplam sanilip siliniyor ve beyan toplami o satirin borcuna
    esitleniyordu (olculdu: 134 satir / 48.962,59 yerine 133 / 428,17).
    Tarih ve evrak tasiyan satir artik hicbir kosulda toplam sayilmaz.
    """
    for i in (i_tarih, i_evrak):
        if i is not None and hucre_metni(al(i)) is not None:
            return False
    aciklama = hucre_metni(al(i_aciklama)) if i_aciklama is not None else None
    if aciklama is not None and not _RE_TOPLAM_ETIKETI.match(_fold(aciklama)):
        return False
    return any(
        _RE_TOPLAM_ETIKETI.match(_fold(h)) for h in satir if isinstance(h, str) and h.strip()
    )


def _kisi_gecerli_mi(isim: str | None) -> bool:
    """Metnin gercekten bir kisi adi olup olmadigini kabaca dogrular.

    En az iki token, her token en az iki harf, hicbir token hizmet kelimesi
    olmamali ve toplam uzunluk 5 harften kisa olmamali.
    """
    if not isim:
        return False
    tokenlar = isim.split()
    if len(tokenlar) < 2:
        return False
    if any(len(t) < 2 for t in tokenlar):
        return False
    if any(t in _HIZMET_KELIMELERI for t in tokenlar):
        return False
    return sum(len(t) for t in tokenlar) >= 5


def _temizle_ve_dogrula(ham: str) -> str | None:
    """Ham metni temizler, kuyrugunu kirpar ve gecerliyse dondurur."""
    isim = _kuyruk_kirp(kisi_metnini_temizle(ham))
    return isim if _kisi_gecerli_mi(isim) else None


# --------------------------------------------------------------------------
# Gider tipine gore kisi adi cikarma
# --------------------------------------------------------------------------

def _bilet_kisi(aciklama: str) -> str | None:
    """Bilet satirindan yolcu adini cikarir.

    Desteklenen bicimler:
        'TK1234567890 SAHTEOGLU/MUSTAFAKEMAL MR  IST-CDG BILET BEDELI'
        'PC1234567891 DENEMECI MEHMET\\M  KYA-SAW-LED BILET BEDELI'
        'VF1234567892 YAPAYSOY ARAS\\I  VKO-SAW-ESB-VKO BILET BEDELI'

    Her ikisinde de SOYAD once gelir; token sirasi korunur (eslestirici
    sirasiz token kumesi kullanir).
    """
    metin = _fold(aciklama)
    if not metin:
        return None
    onek = _RE_BILET_ONEK.match(metin)
    kalan = onek.group("kalan") if onek else metin

    # Isim, asagidaki isaretlerden en erken gelenine kadar surer.
    kesimler: list[int] = []
    for desen in (_RE_YOLCU_ISARETI, _RE_UNVAN, _RE_IATA_ZINCIR):
        esle = desen.search(kalan)
        if esle is not None and esle.start() > 0:
            kesimler.append(esle.start())
    ham_isim = kalan[: min(kesimler)] if kesimler else kalan
    return _temizle_ve_dogrula(ham_isim)


def _otel_kisi(aciklama: str) -> list[str]:
    """Otel satirindan konaklayan kisi adlarini cikarir.

    Kisi adi BASTA, 'AD SOYAD' sirasiyla gelir; ardindan otel adi ve
    '[giris] - [cikis] (gece) KONAKLAMA ...' kuyrugu bulunur. Bir satirda
    ';' ile ayrilmis birden fazla kisi olabilir.

    >>> _otel_kisi("MUSTAFA KEMAL SAHTEOGLU ; POLINA PRIMEROVA GRAND HYATT ISTANBUL [10.07.2026] - [11.07.2026]  (1) KONAKLAMA YURTICI")
    ['MUSTAFA KEMAL SAHTEOGLU', 'POLINA PRIMEROVA']
    """
    metin = _fold(aciklama)
    if not metin:
        return []

    # Tarih blogundan onceki kisim kisi + otel adini icerir.
    kose = metin.find("[")
    if kose > 0:
        bas = metin[:kose]
    else:
        tarih = _RE_TARIH.search(metin)
        bas = metin[: tarih.start()] if tarih is not None and tarih.start() > 0 else metin

    isimler: list[str] = []
    for parca in _RE_KISI_AYIRAC.split(bas):
        parca = parca.strip()
        if not parca:
            continue
        otel = _RE_OTEL_ANAHTAR.search(parca)
        anahtar_var = otel is not None and otel.start() > 0
        isim = _temizle_ve_dogrula(parca)
        if isim is None:
            continue
        tokenlar = isim.split()
        if not anahtar_var and len(tokenlar) > 4:
            # Otel adi anahtar kelime icermiyor; ilk uc tokeni isim kabul et.
            isim = " ".join(tokenlar[:3])
        if isim not in isimler:
            isimler.append(isim)
    return isimler


def _vize_kisi(aciklama: str) -> str | None:
    """Vize satirindan kisi adini cikarir.

    >>> _vize_kisi("HALIT CAN KARADUMAN RUSYA FEDERASYONU TURISTIK E-VIZE")
    'HALIT CAN KARADUMAN'
    """
    return _temizle_ve_dogrula(_fold(aciklama))


def _bagaj_kisi(aciklama: str) -> str | None:
    """Ekstra bagaj satirindan kisi adini cikarir.

    Kisi adi 'TARAFINDAN' kelimesinden once, hizmet aciklamasinin ORTASINDA
    bulunur:
        'EKSTRA BAGAJ UCRETI ISA MUCAHIT SAHIN TARAFINDAN TASINDI'
    """
    metin = _fold(aciklama)
    esle = _RE_BAGAJ.search(metin)
    if esle is not None:
        isim = _temizle_ve_dogrula(esle.group("isim"))
        if isim is not None:
            return isim
    return _temizle_ve_dogrula(metin)


def _diger_kisi(aciklama: str) -> str | None:
    """Transfer / diger hizmet satirlarindan kisi adini cikarir.

    Bu satirlarin bir kismi kisi icermez ('[13.07.2026] - [14.07.2026]
    CENAZE CELENGI'); o durumda None doner.
    """
    metin = _fold(aciklama).strip()
    if not metin or metin.startswith("["):
        return None
    kesimler = [len(metin)]
    kose = metin.find("[")
    if kose > 0:
        kesimler.append(kose)
    ok = metin.find(">")
    if ok > 0:
        kesimler.append(ok)
    tarih = _RE_TARIH.search(metin)
    if tarih is not None and tarih.start() > 0:
        kesimler.append(tarih.start())
    return _temizle_ve_dogrula(metin[: min(kesimler)])


def _paylasim_etiketi_taniniyor_mu(ad: str) -> bool:
    """Paylasim etiketi bilinen bir tuzel kisi mi ('RHI', 'RENSTROYDETAL', 'UST LUGA')?

    Kelimeler bastan sona 1-3 kelimelik tuzel kisi etiketleriyle ortulmeli
    (masraf_merkezi.MasrafMerkeziHaritasi.tuzel_kisi_mi ile ayni kural).
    'GIRIS', 'ODA', 'GECE' gibi hizmet/not kelimeleri paylasim uretmez.
    """
    try:
        from masraf.masraf_merkezi import SIRKET_KANONIK, TUZEL_KISI_ETIKETLERI

        etiketler = set(TUZEL_KISI_ETIKETLERI) | set(SIRKET_KANONIK)
    except Exception:  # noqa: BLE001
        etiketler = {"RHI", "RSD", "RSS", "RC", "RENSTROYDETAL", "RENSERVIS", "UST LUGA"}
    parcalar = [parca for parca in ad.replace("+", " ").split() if parca.isalpha()]
    if not parcalar:
        return False
    i = 0
    while i < len(parcalar):
        uzunluk = 0
        for n in (3, 2, 1):
            if " ".join(parcalar[i:i + n]) in etiketler:
                uzunluk = n
                break
        if not uzunluk:
            return False
        i += uzunluk
    return True


def _paylasim_ayristir(*metinler: str | None) -> list[dict[str, Any]]:
    """'RHI 1/3- RENSTROYDETAL 2/3' bicimli paylasimlari ayristirir.

    Dondurulen her oge: {'masraf_merkezi': str, 'pay': int, 'bolen': int,
    'oran': float}. Paylasim bulunamazsa bos liste doner.
    """
    paylar: list[dict[str, Any]] = []
    for metin in metinler:
        if not metin:
            continue
        katlanmis = _fold(metin)
        for esle in _RE_PAYLASIM.finditer(katlanmis):
            ad = esle.group(1).strip()
            pay, bolen = int(esle.group(2)), int(esle.group(3))
            if not ad or bolen == 0:
                continue
            # Makul sinir: 1 <= pay <= bolen, 2 <= bolen <= 12.
            if pay < 1 or pay > bolen or bolen < 2 or bolen > _PAYLASIM_AZAMI_BOLEN:
                continue
            # Kuyrugundaki hizmet kelimelerini at ('SARJ EDILECEK' gibi)
            ad = _kuyruk_kirp(ad) or ad
            # Etiket bilinen bir tuzel kisi degilse ('ODA 2/3 KISI', 'GECE 1/2')
            # paylasim uretilmez; mahsuplasma taninmayan etiketi zaten reddeder,
            # burada elemek yanlis alarm uyarisini da onler.
            if not _paylasim_etiketi_taniniyor_mu(ad):
                continue
            paylar.append(
                {
                    "masraf_merkezi": ad,
                    "pay": pay,
                    "bolen": bolen,
                    "oran": round(pay / bolen, 6),
                }
            )
    return paylar


def _kisi_coz(gider_tipi: str, aciklama: str) -> tuple[str | None, list[str]]:
    """Gider tipine gore dogru cikarma fonksiyonunu cagirir.

    Returns:
        (birincil kisi adi veya None, tum kisi adlari listesi)
    """
    if gider_tipi == "Bilet":
        isim = _bilet_kisi(aciklama)
        return isim, ([isim] if isim else [])
    if gider_tipi == "Otel":
        isimler = _otel_kisi(aciklama)
        return (isimler[0] if isimler else None), isimler
    if gider_tipi == "Vize":
        isim = _vize_kisi(aciklama)
        return isim, ([isim] if isim else [])
    if gider_tipi == "Bagaj":
        isim = _bagaj_kisi(aciklama)
        return isim, ([isim] if isim else [])
    isim = _diger_kisi(aciklama)
    return isim, ([isim] if isim else [])


def _gider_tipi_coz(islem: str | None, aciklama: str) -> str:
    """'Islem' kolonu ve aciklama metninden gider tipini belirler."""
    katlanmis = _fold(aciklama)
    if "EKSTRA BAGAJ" in katlanmis or "BAGAJ UCRETI" in katlanmis:
        return "Bagaj"
    anahtar = " ".join(_fold(islem).lower().split())
    tip = _ISLEM_TIPLERI.get(anahtar)
    if tip is not None:
        return tip
    if "KONAKLAMA" in katlanmis:
        return "Otel"
    if "VIZE" in katlanmis:
        return "Vize"
    if "BILET" in katlanmis:
        return "Bilet"
    return "Diger"


# --------------------------------------------------------------------------
# Antik ham cari hareket dokumu (.xls)
# --------------------------------------------------------------------------

def antik_cari_oku(yol: str | Path) -> list[GiderSatiri]:
    """Antik Travel ham cari hareket dokumunu (.xls) okur.

    Baslik satiri 'Islem Tarihi' metni aranarak DINAMIK bulunur (sabit satir
    numarasi varsayilmaz). Baslik altindaki hesap adi satiri ('ENERGO-USD-...')
    ve dosya sonundaki TOPLAM satiri atlanir.

    Tutar = Borc - Alacak (iade satirlari negatif olur); ham degerler
    ek['borc'] / ek['alacak'] icinde saklanir.
    """
    p = Path(yol)
    calisma: Calisma = calisma_oku(p)
    if not calisma.sayfalar:
        return []
    sayfa_adi = calisma.sayfa_adlari[0]
    satirlar = calisma.satirlar(sayfa_adi)

    baslik_i = baslik_satiri_bul(satirlar, aranan=("Islem Tarihi", "İşlem Tarihi"))
    if baslik_i < 0:
        return []
    harita = kolon_haritasi(satirlar[baslik_i])

    i_tarih = kolon_ara(harita, "islem tarihi", "tarih")
    i_islem = kolon_ara(harita, "islem", icerir=False)
    i_evrak = kolon_ara(harita, "evrak no", "evrak")
    i_aciklama = kolon_ara(harita, "aciklama")
    i_doviz = kolon_ara(harita, "doviz", "para birimi")
    i_borc = kolon_ara(harita, "borc")
    i_alacak = kolon_ara(harita, "alacak")

    if i_aciklama is None:
        return []

    sonuclar: list[GiderSatiri] = []
    # Dokumun kendi TOPLAM satiri (borc kolonu toplami). Okuyucunun bir borc
    # satirini kacirip kacirmadigini mutabakatta 'Faturada Yazan Toplam'
    # olarak ele verir. Alacak (iade) toplami dokumde beyan edilmez.
    beyan_borc: float | None = None
    alacak_toplami = 0.0
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        if dolu_hucre_sayisi(satir) == 0:
            continue

        def al(i: int | None) -> Any:
            if i is None or i >= len(satir):
                return None
            return satir[i]

        aciklama = hucre_metni(al(i_aciklama))
        toplam_satiri = _toplam_satiri_mi(satir, al, i_tarih, i_evrak, i_aciklama)
        if toplam_satiri and beyan_borc is None:
            beyan_borc = hucre_sayisi(al(i_borc))
        if aciklama is None:
            # Hesap adi satiri ('ENERGO-USD-ENERGO-USD') ve TOPLAM satiri
            continue
        if toplam_satiri:
            continue

        islem = hucre_metni(al(i_islem))
        gider_tipi = _gider_tipi_coz(islem, aciklama)
        kisi, kisiler = _kisi_coz(gider_tipi, aciklama)

        borc = hucre_sayisi(al(i_borc))
        alacak = hucre_sayisi(al(i_alacak))
        alacak_toplami += alacak or 0.0
        tutar: float | None
        if borc is None and alacak is None:
            tutar = None
        else:
            tutar = (borc or 0.0) - (alacak or 0.0)

        guzergah_esle = _RE_IATA_ZINCIR.search(_fold(aciklama))
        ek: dict[str, Any] = {
            "sayfa": sayfa_adi,
            "islem": islem,
            "evrak_no": hucre_metni(al(i_evrak)),
            "borc": borc,
            "alacak": alacak,
            "kisiler": kisiler,
            "kisi_sayisi": len(kisiler),
        }
        if guzergah_esle is not None:
            ek["guzergah"] = guzergah_esle.group(0)

        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="antik_cari",
                satir_no=r + 1,
                belge_tarihi=hucre_tarihi(al(i_tarih), calisma.datemode),
                aciklama=aciklama,
                kisi_ham=kisi,
                sicil_ham=None,
                tckn_ham=None,
                tutar=tutar,
                para_birimi=hucre_metni(al(i_doviz)),
                masraf_merkezi_kaynak=None,  # bu dosyada masraf merkezi YOKTUR
                gider_tipi=gider_tipi,
                ek=ek,
            )
        )
    if beyan_borc is not None and sonuclar:
        # Beyan = dokumun TOPLAM'i (borc) - okunan alacak. Okunan net ile farki,
        # tam olarak 'kacirilan borc satiri' demektir; alacak tarafi dokumde
        # beyan edilmedigi icin dogrulanamaz, bu bilerek boyle.
        ozet = {"TOPLAM (borc)": round(beyan_borc, 2),
                "okunan alacak": -round(alacak_toplami, 2)}
        for s in sonuclar:
            s.ek["fatura_ozeti"] = dict(ozet)
            s.ek["beyan_yontemi"] = "dokumun TOPLAM satiri (borc) - okunan alacak"
    return sonuclar


# --------------------------------------------------------------------------
# Yuzyil elle dagitilmis dosya (.xlsx) - dogruluk referansi
# --------------------------------------------------------------------------

# Ozet satirlarini tanimlayan aciklama metinleri.
_OZET_METINLERI: tuple[str, ...] = (
    "TOPLAM TUTAR", "ODENECEK TOPLAM TUTAR", "ODENECEK", "IADE",
    "GENEL TOPLAM", "TOPLAM",
)


def yuzyil_dagitilmis_oku(yol: str | Path) -> list[GiderSatiri]:
    """Yuzyil elle dagitilmis dosyasini (.xlsx) okur.

    Bu dosya otomasyonun DOGRULUGUNU olcmek icin kullanilir: 'SANTIYESI'
    kolonu insan tarafindan doldurulmustur ve masraf_merkezi_kaynak alanina
    yazilir. 'RHI 1/3- RENSTROYDETAL 2/3' gibi paylasimlar ek['paylasim']
    icine ayristirilir; paylasim notu santiye kolonunda da, yanindaki
    basliksiz not kolonlarinda da bulunabilir.

    Sondaki TOPLAM / IADE / ODENECEK ozet satirlari atlanir.
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    if not calisma.sayfalar:
        return []
    sayfa_adi = calisma.sayfa_adlari[0]
    satirlar = calisma.satirlar(sayfa_adi)

    baslik_i = baslik_satiri_bul(satirlar, aranan=("S.NO", "ACIKLAMA", "AÇIKLAMA"))
    if baslik_i < 0:
        return []
    baslik = satirlar[baslik_i]
    harita = kolon_haritasi(baslik)

    i_sno = kolon_ara(harita, "s no", "sno", "sira no")
    i_tarih = kolon_ara(harita, "fatura tarihi", "tarih")
    i_aciklama = kolon_ara(harita, "aciklama")
    i_tutar = kolon_ara(harita, "satis", "tutar", "toplam")
    i_guzergah = kolon_ara(harita, "ucus guzergahi", "guzergah")
    i_ucus = kolon_ara(harita, "ucus bilgisi", "havayolu")
    i_santiye = kolon_ara(harita, "santiyesi", "santiye", "masraf merkezi")

    if i_aciklama is None:
        return []
    # Baslik satirinda adi olmayan kolonlar not kolonlaridir.
    not_indeksleri = [
        i
        for i in range(len(baslik))
        if hucre_metni(baslik[i]) is None and i not in {i_santiye, i_aciklama}
    ]

    sonuclar: list[GiderSatiri] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        if dolu_hucre_sayisi(satir) == 0:
            continue

        def al(i: int | None) -> Any:
            if i is None or i >= len(satir):
                return None
            return satir[i]

        aciklama = hucre_metni(al(i_aciklama))
        if aciklama is None:
            continue
        katlanmis_aciklama = _fold(aciklama)
        if any(katlanmis_aciklama.startswith(o) for o in _OZET_METINLERI):
            continue
        # Ozet satirlarinda S.NO bostur; veri satirlarinda her zaman doludur.
        if i_sno is not None and hucre_metni(al(i_sno)) is None:
            continue

        guzergah = hucre_metni(al(i_guzergah))
        ucus = hucre_metni(al(i_ucus))
        if guzergah or ucus:
            gider_tipi = "Bilet"
        else:
            gider_tipi = _gider_tipi_coz(None, aciklama)

        if gider_tipi == "Bilet":
            # Bu dosyada bilet satirlarinin aciklamasi ZATEN duz kisi adidir.
            isim = _temizle_ve_dogrula(katlanmis_aciklama)
            kisi, kisiler = isim, ([isim] if isim else [])
        else:
            kisi, kisiler = _kisi_coz(gider_tipi, aciklama)

        santiye = hucre_metni(al(i_santiye))
        notlar = [hucre_metni(al(i)) for i in not_indeksleri]
        notlar = [n for n in notlar if n]
        paylasim = _paylasim_ayristir(santiye, *notlar)

        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="yuzyil_dagitilmis",
                satir_no=r + 1,
                belge_tarihi=hucre_tarihi(al(i_tarih), calisma.datemode),
                aciklama=aciklama,
                kisi_ham=kisi,
                sicil_ham=None,
                tckn_ham=None,
                tutar=hucre_sayisi(al(i_tutar)),
                para_birimi="USD",
                masraf_merkezi_kaynak=santiye,
                gider_tipi=gider_tipi,
                ek={
                    "sayfa": sayfa_adi,
                    "s_no": hucre_metni(al(i_sno)),
                    "guzergah": guzergah,
                    "ucus_bilgisi": ucus,
                    "kisiler": kisiler,
                    "kisi_sayisi": len(kisiler),
                    "notlar": notlar,
                    "paylasim": paylasim,
                },
            )
        )
    return sonuclar
