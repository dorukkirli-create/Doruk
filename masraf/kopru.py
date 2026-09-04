"""TC kimlik numarasi ile sicil arasinda kopru kurar.

Neden gerekli: personel ana verisinde TC kimlik numarasi YOKTUR. Sadece sicil,
ad soyad ve dogum tarihi vardir. Buna karsilik saglik kontrol listesi ve
arabuluculuk dosyasi gibi yardimci kaynaklarda TC kimlik numarasi vardir ama
sicil yoktur. Ikisini dogrudan birlestirmek mumkun degil.

Cozum: ad soyad ve dogum tarihi ortak alandir. Bu ikisi birlikte neredeyse
benzersizdir. Olculen deger: personel verisinde isim cakismasi yuzde 3,6 iken
isim artı dogum tarihinde yuzde 1,68'e duser.

Olculen sonuc (Temmuz 2026 saglik kontrol listesi, 50 kisi):
    isimle personel verisinde bulunan      30
    isim + dogum tarihi ile tek adaya inen 27
    personel verisinde hic olmayan         20

Yani tek calistirmada 27 kayitlik kopru uretilir. Kopru kalici olarak
`veri/tckn_sicil.csv` dosyasina yazilir ve her ay yeni liste geldikce buyur.

Kullanim:
    from masraf.kopru import kopru_turet, kopruyu_deftere_yaz
    adaylar = kopru_turet(satirlar, defter)
    n = kopruyu_deftere_yaz(adaylar, defterler)
"""

from __future__ import annotations

import datetime
import logging
from dataclasses import dataclass
from typing import Iterable

from masraf.metin import isim_normalize, isim_tokenlari

_log = logging.getLogger(__name__)


@dataclass
class KopruAdayi:
    """TC kimlik ile sicil arasinda kurulan tek bir baglanti."""

    tckn: str
    sicil: str
    ad_soyad_kaynak: str
    ad_soyad_personel: str | None
    dogum_tarihi: datetime.date | None
    kaynak_tip: str
    yontem: str  # 'isim_dogum' | 'isim_tek_aday'
    guven: float

    @property
    def aciklama(self) -> str:
        if self.yontem == "isim_dogum":
            g = self.dogum_tarihi.strftime("%d.%m.%Y") if self.dogum_tarihi else "?"
            return (f"Ad soyad ve dogum tarihi ({g}) eslesti: "
                    f"{self.ad_soyad_personel} / {self.sicil}")
        return f"Ad soyad tek adayla eslesti: {self.ad_soyad_personel} / {self.sicil}"


def _tarihe_cevir(v) -> datetime.date | None:
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def _adaylari_bul(defter, ad: str) -> list[str]:
    """Isim icin sicil adaylarini bulur; duz, ters ve token sirasiz dener."""
    norm = isim_normalize(ad)
    if not norm:
        return []
    adaylar = defter.isimle_adaylar(norm)
    if adaylar:
        return adaylar
    # 'AD SOYAD' ile 'SOYAD AD' arasindaki sira farki
    ters = " ".join(reversed(norm.split()))
    adaylar = defter.isimle_adaylar(ters)
    if adaylar:
        return adaylar
    # token kumesi (patronimik gibi fazladan tokenlari tolere eder)
    return defter.token_ile_adaylar(isim_tokenlari(norm))


def kopru_turet(
    satirlar: Iterable,
    defter,
    dogum_zorunlu: bool = True,
) -> list[KopruAdayi]:
    """TC kimlik tasiyan satirlari personel defterindeki sicillere baglar.

    Args:
        satirlar: GiderSatiri listesi. `tckn_ham` dolu olanlar degerlendirilir.
            Dogum tarihi `ek['dogum_tarihi']` icinde aranir.
        defter: PersonelDefteri ornegi.
        dogum_zorunlu: True ise sadece dogum tarihi de eslesen kayitlar
            dondurulur. False ise isim tek aday verdiginde dogum tarihi
            olmasa da kabul edilir, ama guven skoru dusuktur.

    Returns:
        Kurulabilen kopru adaylari. Ayni TC birden fazla kez gelirse
        tek kayit dondurulur.
    """
    gorulen: dict[str, KopruAdayi] = {}
    for s in satirlar:
        tckn = (getattr(s, "tckn_ham", None) or "").strip()
        ad = getattr(s, "kisi_ham", None)
        if not tckn or not ad or tckn in gorulen:
            continue

        adaylar = _adaylari_bul(defter, ad)
        if not adaylar:
            continue

        ek = getattr(s, "ek", None) or {}
        dogum = _tarihe_cevir(ek.get("dogum_tarihi"))

        secilen: str | None = None
        yontem = ""
        guven = 0.0

        if dogum is not None:
            uyan = []
            for sic in adaylar:
                kayit = defter.sicil_ile(sic) or {}
                if _tarihe_cevir(kayit.get("dogum_tarihi")) == dogum:
                    uyan.append(sic)
            if len(uyan) == 1:
                secilen, yontem, guven = uyan[0], "isim_dogum", 0.97
            elif len(uyan) > 1:
                # ayni isim ayni dogum tarihi: karar verilemez, atla
                _log.warning("TC %s icin %d aday ayni dogum tarihini tasiyor, atlandi",
                             tckn, len(uyan))
                continue

        if secilen is None and not dogum_zorunlu and len(adaylar) == 1:
            secilen, yontem, guven = adaylar[0], "isim_tek_aday", 0.70

        if secilen is None:
            continue

        kayit = defter.sicil_ile(secilen) or {}
        gorulen[tckn] = KopruAdayi(
            tckn=tckn,
            sicil=secilen,
            ad_soyad_kaynak=str(ad),
            ad_soyad_personel=kayit.get("ad_soyad"),
            dogum_tarihi=dogum,
            kaynak_tip=getattr(s, "kaynak_tip", "?"),
            yontem=yontem,
            guven=guven,
        )
    return list(gorulen.values())


def kopruyu_deftere_yaz(adaylar: Iterable[KopruAdayi], defterler) -> int:
    """Kopru adaylarini Defterler.tckn_sicil tablosuna yazar.

    Zaten kayitli olan TC numaralari atlanir; mevcut kayit ustune yazilmaz
    cunku elle onaylanmis bir kayit turetilmis bir kayittan daha guvenilirdir.

    Returns:
        Yeni eklenen kayit sayisi.
    """
    eklenen = 0
    for a in adaylar:
        if a.tckn in getattr(defterler, "tckn_sicil", {}):
            continue
        try:
            defterler.tckn_kopru_ekle(
                a.tckn, a.sicil,
                ad_soyad=a.ad_soyad_personel or a.ad_soyad_kaynak,
                kaynak=f"turetildi:{a.yontem}",
            )
            eklenen += 1
        except TypeError:
            # imza farkliysa en sade cagriyi dene
            defterler.tckn_kopru_ekle(a.tckn, a.sicil)
            eklenen += 1
        except Exception as e:
            _log.warning("Kopru yazilamadi TC %s: %s", a.tckn, e)
    if eklenen:
        try:
            defterler.kaydet()
        except Exception as e:
            _log.warning("Defter kaydedilemedi: %s", e)
    return eklenen


def kopru_ozeti(satirlar: Iterable, defter, adaylar: list[KopruAdayi]) -> dict:
    """Kopru kurma isleminin ozetini dondurur; arayuzde gosterilir."""
    satirlar = list(satirlar)
    tckn_tasiyan = {(getattr(s, "tckn_ham", None) or "").strip()
                    for s in satirlar if getattr(s, "tckn_ham", None)}
    tckn_tasiyan.discard("")
    isimle_bulunan = 0
    for t in tckn_tasiyan:
        ilgili = next((s for s in satirlar
                       if (getattr(s, "tckn_ham", None) or "").strip() == t), None)
        if ilgili and _adaylari_bul(defter, getattr(ilgili, "kisi_ham", "") or ""):
            isimle_bulunan += 1
    return {
        "tckn_tasiyan_kisi": len(tckn_tasiyan),
        "isimle_bulunan": isimle_bulunan,
        "kopru_kurulan": len(adaylar),
        "personel_verisinde_yok": len(tckn_tasiyan) - isimle_bulunan,
    }


# --------------------------------------------------------------------------
# Dogum tarihi ile dogrulanmis alias turetme
# --------------------------------------------------------------------------
#
# Neden gerekli: bulanik ve aile kademeleri isim benzerligine bakar ve
# yanilabilir. Olculen ornek (Temmuz 2026 saglik kontrol listesi):
#
#   GOKHAN GUZEL   (dogum 18.04.1980) -> Guzel Serdal        (15.07.1977)  YANLIS
#   MEHMET E NERGIZ(dogum 05.06.1989) -> Nergiz Mehmet Kerem (08.09.1981)  YANLIS
#   BARIS GOCEDEN  (dogum 20.02.1984) -> Baris Baris         (10.06.1986)  YANLIS
#   OGUN BIZ       (dogum 06.08.1992) -> Biz Selami          (15.02.1991)  YANLIS
#   SEHRIBAN OZKAN (dogum 03.01.1995) -> Seriban Ozkan       (03.01.1995)  DOGRU
#   YASAR MERT YANAR(dogum 17.10.1994)-> Banar Yasar Mert    (17.10.1994)  DOGRU
#   MEHMET DALKILIC(dogum 24.04.1987) -> Dalkilinc Mehmet    (24.04.1987)  DOGRU
#
# Dogum tarihi dortunu de eler, ucunu de dogrular. Yazim farki (Yanar/Banar,
# Sehriban/Seriban, Dalkilic/Dalkilinc) dogum tarihi ile guvenle asilir.

_BENZERLIK_ESIGI = 78  # rapidfuzz token_set_ratio alt siniri


def _dogum_ile_daralt(defter, adaylar: Iterable[str], dogum: datetime.date) -> list[str]:
    """Aday siciller icinden dogum tarihi birebir tutanlari secer."""
    uyan = []
    for sic in adaylar:
        kayit = defter.sicil_ile(sic) or {}
        if _tarihe_cevir(kayit.get("dogum_tarihi")) == dogum:
            uyan.append(sic)
    return uyan


def _tum_siciller(defter) -> list[str]:
    """Defterdeki isimli tum sicilleri dondurur.

    PersonelDefteri herkese acik bir "hepsini gez" metodu sunmuyor. Isim
    indeksi tum isimli sicilleri kapsadigi icin oradan turetiyoruz; indeks
    yoksa soyad indeksine, o da yoksa bos listeye duseriz.
    """
    for alan in ("_isim_index", "_soyad_index", "_token_index"):
        indeks = getattr(defter, alan, None)
        if isinstance(indeks, dict) and indeks:
            gorulen: set[str] = set()
            for siciller in indeks.values():
                gorulen.update(siciller)
            return list(gorulen)
    return []


def _dogum_indeksi(defter) -> dict[datetime.date, list[str]]:
    """Dogum tarihi -> sicil listesi indeksini bir kez kurar ve onbellekler."""
    onbellek = getattr(defter, "_kopru_dogum_indeksi", None)
    if onbellek is not None:
        return onbellek
    indeks: dict[datetime.date, list[str]] = {}
    for sicil in _tum_siciller(defter):
        kayit = defter.sicil_ile(sicil) or {}
        dog = _tarihe_cevir(kayit.get("dogum_tarihi"))
        if dog is not None:
            indeks.setdefault(dog, []).append(sicil)
    try:
        setattr(defter, "_kopru_dogum_indeksi", indeks)
    except Exception:  # pragma: no cover
        pass
    return indeks


def alias_turet(satirlar: Iterable, defter, esik: int = _BENZERLIK_ESIGI) -> list[KopruAdayi]:
    """Dogum tarihi ile dogrulanmis isim eslesmeleri uretir.

    Once isim tokenlarindan bir aday havuzu kurulur, sonra SADECE dogum tarihi
    birebir tutan adaylar birakilir, en sonunda isim benzerligi esigi uygulanir.
    Dogum tarihi olmayan satirlar atlanir; bu fonksiyonun tum guvenligi o
    alandan gelir.

    Args:
        satirlar: GiderSatiri listesi. `ek['dogum_tarihi']` dolu olmali.
        defter: PersonelDefteri ornegi.
        esik: rapidfuzz token_set_ratio alt siniri.

    Returns:
        Dogrulanmis eslesmeler. `yontem` degeri 'isim_dogum_bulanik'.
    """
    try:
        from rapidfuzz import fuzz
    except ImportError:  # pragma: no cover
        _log.warning("rapidfuzz kurulu degil, alias turetme atlandi")
        return []

    sonuc: dict[str, KopruAdayi] = {}
    for s in satirlar:
        ad = getattr(s, "kisi_ham", None)
        ek = getattr(s, "ek", None) or {}
        dogum = _tarihe_cevir(ek.get("dogum_tarihi"))
        if not ad or dogum is None:
            continue
        norm = isim_normalize(ad)
        if not norm or norm in sonuc:
            continue
        # zaten kesin eslesenler icin alias uretmeye gerek yok
        if defter.isimle_adaylar(norm):
            continue

        uyan = _dogum_indeksi(defter).get(dogum) or []
        if not uyan:
            continue

        puanlar = []
        for sic in uyan:
            kayit = defter.sicil_ile(sic) or {}
            hedef = kayit.get("ad_soyad_norm") or isim_normalize(kayit.get("ad_soyad") or "")
            if hedef:
                puanlar.append((fuzz.token_set_ratio(norm, hedef), sic, kayit))
        if not puanlar:
            continue
        puanlar.sort(reverse=True, key=lambda x: x[0])
        en_iyi = puanlar[0]
        if en_iyi[0] < esik:
            continue
        # ayni dogum tarihinde iki yakin isim varsa karar verme
        if len(puanlar) > 1 and puanlar[1][0] >= en_iyi[0] - 4:
            _log.warning("%s icin ayni dogum tarihinde iki yakin aday var, atlandi", norm)
            continue

        sonuc[norm] = KopruAdayi(
            tckn=(getattr(s, "tckn_ham", None) or "").strip(),
            sicil=en_iyi[1],
            ad_soyad_kaynak=str(ad),
            ad_soyad_personel=en_iyi[2].get("ad_soyad"),
            dogum_tarihi=dogum,
            kaynak_tip=getattr(s, "kaynak_tip", "?"),
            yontem="isim_dogum_bulanik",
            guven=0.93,
        )
    return list(sonuc.values())


def aliaslari_deftere_yaz(adaylar: Iterable[KopruAdayi], defterler) -> int:
    """Turetilmis aliaslari Defterler.aliases tablosuna yazar.

    Elle onaylanmis mevcut kayitlarin ustune yazmaz.
    """
    eklenen = 0
    for a in adaylar:
        norm = isim_normalize(a.ad_soyad_kaynak)
        if not norm or norm in getattr(defterler, "aliases", {}):
            continue
        try:
            defterler.alias_ekle(norm, a.sicil,
                                 ad_soyad=a.ad_soyad_personel or a.ad_soyad_kaynak,
                                 kaynak=f"turetildi:{a.yontem}")
            eklenen += 1
        except TypeError:
            defterler.alias_ekle(norm, a.sicil)
            eklenen += 1
        except Exception as e:
            _log.warning("Alias yazilamadi %s: %s", norm, e)
    if eklenen:
        try:
            defterler.kaydet()
        except Exception as e:
            _log.warning("Defter kaydedilemedi: %s", e)
    return eklenen
