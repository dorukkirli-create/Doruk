"""Masraf merkezi haritasina eklenecek satirlari personel verisinden turetir.

Neden gerekli: bir gorev yeri haritada tanimli degilse kod onu metin olarak
tasir ve ciktida isaretler. Operatorun o satiri elle arastirip hangi sirkete
ait oldugunu bulmasi gerekir. Oysa bilgi zaten elimizdedir.

1C personel listesindeki ``Firm 2`` kolonu her projenin tuzel kisisini
verir ve haritadaki ``sirket`` kolonuyla AYNI sozlugu kullanir. Olculdu:
haritada tanimli projelerin tamaminda iki kaynak birebir ortusuyor
(GPP Project -> UST LUGA, Udokan (GMK) -> RHI, ...).

Bu modul o baglantiyi kurar: tanimsiz her gorev yeri icin kac kisinin o
projede calistigini, tuzel kisisini ve haritaya yapistirilmaya hazir bir
satiri uretir.

Onerilen KOD bir tahmindir. Finans kendi kodunu kullanmalidir; kod yalnizca
satirin bos kalmamasi ve hemen calisabilmesi icin uretilir. Bu yuzden her
oneri ``kod_onerisi`` bayragiyla isaretlenir ve ciktida "finans onaylamali"
diye gosterilir.
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass
from typing import Any, Iterable, Sequence

__all__ = ["HaritaOnerisi", "oneri_uret", "ONERI_BASLIKLARI"]

#: Cikti sayfasinin kolonlari.
ONERI_BASLIKLARI: tuple[tuple[str, str, int], ...] = (
    ("Gorev Yeri (personel verisindeki hali)", "metin", 40),
    ("Onerilen Masraf Merkezi Kodu", "metin", 26),
    ("Masraf Merkezi Adi", "metin", 34),
    ("Sirket (1C Firm 2)", "metin", 18),
    ("Bu projedeki kisi", "tamsayi", 16),
    ("Faturadaki satir", "tamsayi", 15),
    ("Faturadaki tutar", "sayi", 16),
    ("Kaynak", "metin", 22),
)

#: Koddan atilacak, ayirt edici olmayan kelimeler.
_DOLGU = frozenset({
    "PROJECT", "PROJESI", "PROJE", "THE", "AND", "VE", "OF", "FOR",
    "SERVICES", "SERVICE", "OOO", "LLC", "AS", "A.S",
})

#: Sik gecen uzun kelimelerin kisaltmalari. Kod okunakli kalsin diye.
_KISALT = {
    "MANAGEMENT": "MGMT", "BUSINESS": "BUS", "CENTER": "CTR", "CENTRE": "CTR",
    "PRODUCTION": "PROD", "RENSTROYDETAL": "RSD", "RENSERVIS": "RSS",
    "CATERING": "CATER", "MURMANSK": "MRM", "NOVOSIBIRSK": "NSK",
    "HEADQUARTER": "HQ", "TECHNICAL": "TECH", "OFFICE": "OFC",
}


@dataclass
class HaritaOnerisi:
    """Haritaya eklenmeye hazir tek bir satir onerisi."""

    gorev_yeri: str
    kod: str
    ad: str
    sirket: str
    kisi_sayisi: int = 0
    satir_sayisi: int = 0
    tutar: float = 0.0
    para_birimi: str = ""
    kaynak: str = "1C Firm 2"
    kod_onerisi: bool = True

    def csv_satiri(self) -> str:
        """masraf_merkezi_haritasi.csv dosyasina yapistirilacak satir."""
        return f"{self.gorev_yeri},{self.kod},{self.ad},{self.sirket},E"


def _katla(metin: str) -> str:
    """Turkce ve Kiril harfleri ASCII'ye indirir, buyuk harfe cevirir."""
    esle = {"ı": "i", "İ": "I", "ş": "s", "Ş": "S", "ğ": "g", "Ğ": "G",
            "ü": "u", "Ü": "U", "ö": "o", "Ö": "O", "ç": "c", "Ç": "C"}
    for a, b in esle.items():
        metin = metin.replace(a, b)
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(c for c in metin if not unicodedata.combining(c)).upper()


def kod_uret(gorev_yeri: str, kullanilan: set[str] | None = None) -> str:
    """Gorev yerinden okunakli, benzersiz bir kod onerir.

    'Renstroydetal - Ust-Luga GPC' -> 'RSD-UST-LUGA-GPC'
    'Bsk Management Group'         -> 'BSK-MGMT-GROUP'

    Kod finansin onayina tabidir; buradaki amac makul ve okunabilir bir
    baslangic degeri uretmektir.
    """
    kullanilan = kullanilan if kullanilan is not None else set()
    kelimeler = [k for k in re.split(r"[^A-Za-z0-9]+", _katla(gorev_yeri)) if k]
    secilen: list[str] = []
    for k in kelimeler:
        if k in _DOLGU:
            continue
        secilen.append(_KISALT.get(k, k))
    if not secilen:
        secilen = kelimeler[:1] or ["MERKEZ"]
    kod = "-".join(secilen[:4])[:24].strip("-")
    if kod not in kullanilan:
        return kod
    for i in range(2, 60):
        aday = f"{kod[:21]}-{i}"
        if aday not in kullanilan:
            return aday
    return kod


def _sirketi_bul(gorev_yeri: str, defterler: Sequence[Any]) -> tuple[str, int]:
    """Bir gorev yerinde calisanlarin tuzel kisisi ve kisi sayisi.

    1C listesindeki 'Firm 2' (kayitta ``sirket2``) tercih edilir; yoksa
    ``sirket`` kolonuna duser. En sik gecen deger secilir.
    """
    hedef = " ".join(_katla(gorev_yeri).split())
    sayac: Counter[str] = Counter()
    toplam = 0
    for defter in defterler:
        if defter is None:
            continue
        kayitlar = getattr(defter, "_kayitlar", None) or getattr(defter, "_sicil", {}).values()
        for kayit in kayitlar:
            yer = kayit.get("gorev_yeri")
            if not yer or " ".join(_katla(str(yer)).split()) != hedef:
                continue
            toplam += 1
            firma = (kayit.get("sirket2") or kayit.get("sirket") or "").strip()
            if firma:
                sayac[firma] += 1
    if not sayac:
        return "", toplam
    return sayac.most_common(1)[0][0], toplam


def oneri_uret(
    sonuclar: Iterable[Any],
    harita: Any,
    defter: Any = None,
    yardimci: Any = None,
) -> list[HaritaOnerisi]:
    """Ciktida 'haritada tanimli degil' cikan gorev yerleri icin oneri uretir.

    Args:
        sonuclar: ``Sonuc`` listesi. Hangi gorev yerinin kac satir ve ne kadar
            tutar tasidigini buradan olceriz; oncelik siralamasi icin gerekli.
        harita: ``MasrafMerkeziHaritasi``. Zaten tanimli olanlar elenir.
        defter: Ana ``PersonelDefteri`` (istege bagli).
        yardimci: ``YardimciDefter`` (1C listesi). Firm 2 buradan gelir.

    Returns:
        Tutari buyukten kucuge sirali oneri listesi.
    """
    olcum: dict[str, dict] = {}
    for s in sonuclar:
        ek = s.satir.ek if isinstance(s.satir.ek, dict) else {}
        if ek.get("masraf_merkezi_haritada"):
            continue
        yer = s.gorev_yeri
        if not yer:
            continue
        kayit = olcum.setdefault(str(yer), {"satir": 0, "tutar": 0.0, "pb": ""})
        kayit["satir"] += 1
        if s.satir.tutar is not None:
            kayit["tutar"] += float(s.satir.tutar)
            kayit["pb"] = kayit["pb"] or (s.satir.para_birimi or "")

    defterler = [d for d in (yardimci, defter) if d is not None]
    kullanilan = set(harita.kod_adlari()) if hasattr(harita, "kod_adlari") else set()

    oneriler: list[HaritaOnerisi] = []
    for yer, olc in olcum.items():
        if hasattr(harita, "coz") and harita.coz(yer):
            continue  # arada haritaya eklenmis olabilir
        sirket, kisi = _sirketi_bul(yer, defterler)
        kod = kod_uret(yer, kullanilan)
        kullanilan.add(kod)
        oneriler.append(HaritaOnerisi(
            gorev_yeri=yer, kod=kod, ad=yer, sirket=sirket,
            kisi_sayisi=kisi, satir_sayisi=olc["satir"],
            tutar=round(olc["tutar"], 2), para_birimi=olc["pb"],
            kaynak="1C Firm 2" if sirket else "sirket bulunamadi",
        ))
    return sorted(oneriler, key=lambda o: (-o.tutar, o.gorev_yeri))


def oneri_satir_degerleri(oneri: HaritaOnerisi) -> list[Any]:
    """ONERI_BASLIKLARI sirasina cevirir."""
    return [
        oneri.gorev_yeri,
        oneri.kod,
        oneri.ad,
        oneri.sirket or "(bulunamadi)",
        oneri.kisi_sayisi,
        oneri.satir_sayisi,
        oneri.tutar,
        oneri.kaynak,
    ]
