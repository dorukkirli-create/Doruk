"""PDF metin katmanini cikarir ve gorunmez karakterleri temizler.

Neden ayri bir modul: e-Fatura ve e-Arsiv PDF'leri ekranda ayni gorunen ama
byte olarak farkli ayiricilar kullaniyor. Olculdu (Energo Mayis-Haziran
yansitma maili, 17 PDF): naif bir regex kumesi 60 alanin 42'sini buluyor,
metin once normalize edilince 60'inin 60'ini buluyor. Kaybettiren karakter
YUMUSAK TIRE (U+00AD): 'Fatura No: GIB...' icinde ekranda tire gibi gorunur,
regex tutmaz, alan sessizce bos doner. Bu sinifta bir hata kimsenin gozune
carpmaz, o yuzden normalizasyon fonksiyonunun kendi birim testi vardir.

pypdf saf Python'dur, Windows'ta derleme ve calisma zamaninda internet
istemez. Kurulu degilse modul yine ice aktarilir; PYPDF_VAR False olur ve
cagiran taraf 'PDF kutuphanesi kurulu degil' diye durustce raporlar.
"""

from __future__ import annotations

import logging
import unicodedata
from pathlib import Path

_log = logging.getLogger(__name__)

try:  # pragma: no cover - ortama bagli
    import pypdf

    PYPDF_VAR = True
except ImportError:  # pragma: no cover - ortama bagli
    pypdf = None
    PYPDF_VAR = False

__all__ = ["PYPDF_VAR", "metni_normalize", "pdf_metni", "PdfMetni"]

#: Ekranda gorunmeyen ya da duz karsiligiyla ayni gorunen karakterler.
#: Anahtar: PDF'te gecen karakter, deger: regex'in bekledigi duz karsiligi.
_DEGISTIR = {
    "­": "",    # SOFT HYPHEN - Koc ve bir Assessment kopyasinda var
    "​": "",    # ZERO WIDTH SPACE
    "‌": "",    # ZERO WIDTH NON-JOINER
    "‍": "",    # ZERO WIDTH JOINER
    "﻿": "",    # BOM
    " ": " ",   # NO-BREAK SPACE
    " ": " ",   # FIGURE SPACE
    " ": " ",   # NARROW NO-BREAK SPACE
    "–": "-",   # EN DASH
    "—": "-",   # EM DASH
    "−": "-",   # MINUS SIGN
    "‘": "'", "’": "'",
    "“": '"', "”": '"',
}

#: Bu uzunlugun altinda metin cikan PDF taranmis goruntu sayilir. Olculdu:
#: metin katmanli PDF'lerde en kisa cikti 1.233 karakter, taranmislarda 0.
TARANMIS_ESIGI = 50


def metni_normalize(metin: str) -> str:
    """Gorunmez ayiricilari temizler, satir sonlarini ve boslugu sadelestirir.

    Yalnizca ekranda ayni gorunen karakterleri degistirir; harf, rakam ve
    noktalama oldugu gibi kalir. Turkce karakterler korunur (NFKC harfleri
    bozmaz, yalnizca uyumluluk formlarini acar).
    """
    if not metin:
        return ""
    for eski, yeni in _DEGISTIR.items():
        metin = metin.replace(eski, yeni)
    metin = unicodedata.normalize("NFKC", metin)
    # NFKC bazi bosluklari geri getirebilir; ikinci gecis onlari da duzler.
    metin = metin.replace(" ", " ")
    satirlar = [" ".join(s.split()) for s in metin.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    return "\n".join(satirlar)


class PdfMetni:
    """Bir PDF'ten cikarilan metin ve okumanin nasil gittigi.

    ``metin`` her zaman normalize edilmis haldedir. ``taranmis`` True ise
    dosyada metin katmani yoktur (saf goruntu); bu dosya OCR olmadan
    okunamaz ve cagiran taraf bunu kullaniciya soylemek zorundadir.
    """

    __slots__ = ("metin", "sayfa_sayisi", "taranmis", "hata")

    def __init__(self, metin: str = "", sayfa_sayisi: int = 0,
                 taranmis: bool = False, hata: str = "") -> None:
        self.metin = metin
        self.sayfa_sayisi = sayfa_sayisi
        self.taranmis = taranmis
        self.hata = hata

    @property
    def okundu_mu(self) -> bool:
        return bool(self.metin) and not self.taranmis and not self.hata

    def __repr__(self) -> str:  # pragma: no cover - tanilama
        return (f"PdfMetni(karakter={len(self.metin)}, sayfa={self.sayfa_sayisi}, "
                f"taranmis={self.taranmis}, hata={self.hata!r})")


def pdf_metni(yol: str | Path) -> PdfMetni:
    """PDF'i acar, metnini cikarir ve normalize eder. Istisna FIRLATMAZ.

    Sorun ciktiginda bos metin ve dolu bir ``hata`` alani doner: okuyucu
    sozlesmesi geregi (bkz. kesif.PARSERLAR) bir ek yuzunden butun mail
    hataya dusmemeli.
    """
    p = Path(yol)
    if not PYPDF_VAR:
        return PdfMetni(hata="PDF kutuphanesi (pypdf) kurulu degil")
    try:
        okuyucu = pypdf.PdfReader(str(p))
        sayfalar = list(okuyucu.pages)
    except Exception as hata:  # pragma: no cover - bozuk dosya
        _log.debug("PDF acilamadi: %s (%s)", p.name, hata)
        return PdfMetni(hata=f"PDF acilamadi: {type(hata).__name__}")

    parcalar = []
    for sayfa in sayfalar:
        try:
            parcalar.append(sayfa.extract_text() or "")
        except Exception as hata:  # pragma: no cover - bozuk sayfa
            _log.debug("PDF sayfasi okunamadi: %s (%s)", p.name, hata)
    metin = metni_normalize("\n".join(parcalar))
    if len(metin.strip()) < TARANMIS_ESIGI:
        return PdfMetni(metin="", sayfa_sayisi=len(sayfalar), taranmis=True)
    return PdfMetni(metin=metin, sayfa_sayisi=len(sayfalar))
