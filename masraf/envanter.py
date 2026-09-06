"""Dosya envanteri: hangi dosya okundu, hangisi atlandi, neden.

Finans 'bu Excel hangi dosyalardan uretildi, hangi ekler okunmadi' sorusunu
elle kontrol edebilmeli. Bu yuzden boru hatti okudugu, atladigi ve
okuyamadigi HER dosya icin bir kayit tutar; kayit Excel'in 'Dosyalar'
sayfasina, kapaga ve calistirma kaydina (OZET.txt) yazilir.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass
from typing import Any, Iterable

#: Durum etiketleri. ASCII tutulur; Excel ve konsolda oldugu gibi gorunur.
OKUNDU = "OKUNDU"              # gider satiri uretti, dagilima girdi
KUTUK = "KUTUK"                # kisi listesi; defter beslemesine girdi, dagilima girmedi
DETAY_LISTESI = "DETAY LISTESI"  # tutarsiz fatura detay listesi; capraz kontrol edildi
ATLANDI = "ATLANDI"            # tablo degil (PDF, docx...); acilmadi
AYNI_ICERIK = "AYNI ICERIK"    # daha once okunan dosyayla birebir ayni; cift sayim olmasin diye atlandi
OKUNAMADI = "OKUNAMADI"        # acilamadi / hata / parola
SATIR_YOK = "SATIR YOK"        # acildi ama hicbir gider satiri cikmadi
MAIL = "MAIL"                  # Outlook mesaji (kapsayici); ekleri ayri kayitlardir
ARSIV = "ARSIV"                # zip arsivi (kapsayici); icindekiler ayri kayitlardir
PERSONEL = "PERSONEL"          # personel verisine benziyor, fatura olarak islenmedi

#: Dagilima giren durumlar. Digerleri tutar tasimaz ya da tasisa da sayilmaz.
DAGILIMA_GIREN = frozenset({OKUNDU})

_KUTUK_TIPLERI = frozenset({"referans_liste", "energo_saglik", "koc_katilimci"})
#: Disaridan kullanilan adlar. Mahsuplasma detay listelerini de kutuk sayar
#: (tutar tasimazlar); envanter onlari ayri durumla (DETAY LISTESI) gosterir.
KUTUK_TIPLERI = _KUTUK_TIPLERI
DETAY_TIPLERI = frozenset({"energo_assessment_detay"})


@dataclass
class DosyaKaydi:
    """Envanterdeki tek bir dosya (ust duzey dosya ya da mail eki)."""

    ad: str
    kaynak: str = ""            # nereden geldi: 'mail.msg > Arabulucu.zip'
    tur: str = ""               # tespit edilen tip: antik_cari, energo_assessment, pdf...
    durum: str = OKUNDU
    sebep: str = ""             # kullaniciya gosterilecek aciklama
    satir: int = 0              # uretilen satir sayisi
    tutarli_satir: int = 0      # tutari okunan satir sayisi
    tutar: float | None = None  # okunan tutar toplami (tutarli satirlar)
    para_birimi: str | None = None
    boyut: int | None = None    # bayt
    ozet: str | None = None     # sha256 (kisa)

    @property
    def dagilima_girdi(self) -> bool:
        return self.durum in DAGILIMA_GIREN

    def sozluk(self) -> dict[str, Any]:
        d = asdict(self)
        d["dagilima_girdi"] = self.dagilima_girdi
        return d


def satirlardan_kayit(ad: str, kaynak: str, tur: str, satirlar: Iterable[Any],
                      boyut: int | None = None, ozet: str | None = None) -> DosyaKaydi:
    """Okunan satirlardan dosya kaydi uretir; durumu satirlarin tipinden cikarir."""
    satirlar = list(satirlar)
    tipler = Counter(getattr(s, "kaynak_tip", "") for s in satirlar)
    paralar = Counter(getattr(s, "para_birimi", None) for s in satirlar
                      if getattr(s, "para_birimi", None))
    tutarli = [float(s.tutar) for s in satirlar if getattr(s, "tutar", None) is not None]
    if not satirlar:
        durum, sebep = SATIR_YOK, "dosya acildi ama gider satiri cikmadi (kolon adlari taninmamis olabilir)"
    elif tipler and set(tipler) <= {"energo_assessment_detay"}:
        durum, sebep = DETAY_LISTESI, "tutar kolonu yok; kisiler yansitma dosyasiyla capraz kontrol edildi"
    elif tipler and set(tipler) <= _KUTUK_TIPLERI:
        durum, sebep = KUTUK, "kisi listesi; dagilima girmedi (defter beslemesi sonucu uyarilarda)"
    else:
        durum = OKUNDU
        sebep = "" if len(tutarli) == len(satirlar) else f"{len(satirlar) - len(tutarli)} satirda tutar okunamadi"
    return DosyaKaydi(
        ad=ad, kaynak=kaynak, tur=tur or (next(iter(tipler)) if tipler else ""),
        durum=durum, sebep=sebep, satir=len(satirlar), tutarli_satir=len(tutarli),
        tutar=round(sum(tutarli), 2) if tutarli else None,
        para_birimi=(paralar.most_common(1)[0][0] if paralar else None),
        # Tam sha256 saklanir: calistirici bunu ISLENEN_DOSYALAR.txt'ye yazar ki
        # ayni ek sonradan tek basina gelirse 'daha once islendi' denebilsin.
        boyut=boyut, ozet=(ozet or None),
    )


def envanter_ozeti(kayitlar: Iterable[DosyaKaydi]) -> dict[str, int]:
    """Durum bazinda sayim; kapak ve konsol icin."""
    return dict(Counter(k.durum for k in kayitlar))


def _boyut(yol: Any) -> int | None:
    try:
        return int(yol.stat().st_size)
    except (OSError, AttributeError):
        return None
