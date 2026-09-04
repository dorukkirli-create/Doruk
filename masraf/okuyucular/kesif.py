"""Kaynak dosya tipini otomatik tespit eden kesif modulu.

Kullanici dosyayi surukleyip biraktiginda hangi parser'in calistirilacagini
belirler. Tespit tamamen deterministiktir: dosya acilir, SAYFA ADLARI ve ilk
15 satirdaki hucre metinleri ASCII katlanmis bicimde toplanir, ardindan
oncelik sirali ipucu kurallari uygulanir.

Ipuclari:
    'Cari Hareket Dokumu'                -> antik_cari
    'SANTIYESI' + 'UCUS GUZERGAHI'       -> yuzyil_dagitilmis
    'Katilimci' + 'Paket'                -> energo_assessment
    'ARABULUCU'                          -> energo_arabulucu
    'BORDROLU LISTE' | 'TCKN'+'SAGLIK'   -> energo_saglik
    'ID' + 'Alt Fonksiyon'               -> koc_katilimci
    aksi halde                           -> genel
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable

from masraf.modeller import GiderSatiri
from masraf.okuyucular.antik import antik_cari_oku, yuzyil_dagitilmis_oku
from masraf.okuyucular.energo import (
    arabulucu_oku,
    assessment_oku,
    koc_katilimci_oku,
    saglik_oku,
)
from masraf.okuyucular.genel import calisma_oku, genel_oku, kolon_anahtari

__all__ = ["dosya_tipini_bul", "oku", "oku_tip", "PARSERLAR"]

# kaynak_tip -> parser fonksiyonu
PARSERLAR: dict[str, Callable[[str | Path], list[GiderSatiri]]] = {
    "antik_cari": antik_cari_oku,
    "yuzyil_dagitilmis": yuzyil_dagitilmis_oku,
    "energo_assessment": assessment_oku,
    "energo_arabulucu": arabulucu_oku,
    "energo_saglik": saglik_oku,
    "koc_katilimci": koc_katilimci_oku,
    "genel": genel_oku,
}

# Kesif icin okunacak satir sayisi (dosyanin tamami okunmaz, hizli kalir).
_KESIF_SATIR_SINIRI = 15


def _ipuclarini_topla(yol: Path) -> tuple[set[str], str]:
    """Sayfa adlari ve ilk satirlardaki hucre metinlerini toplar.

    Returns:
        (normalize edilmis benzersiz metinler kumesi, hepsinin birlesimi)
    """
    calisma = calisma_oku(yol, satir_siniri=_KESIF_SATIR_SINIRI)
    metinler: set[str] = set()
    for sayfa_adi, satirlar in calisma.sayfalar.items():
        anahtar = kolon_anahtari(sayfa_adi)
        if anahtar:
            metinler.add(anahtar)
        for satir in satirlar:
            for hucre in satir:
                anahtar = kolon_anahtari(hucre)
                if anahtar:
                    metinler.add(anahtar)
    return metinler, " || ".join(sorted(metinler))


def _iceriyor(blob: str, *parcalar: str) -> bool:
    """Birlesik metinde verilen parcalarin HEPSI geciyor mu?"""
    return all(kolon_anahtari(p) in blob for p in parcalar)


def dosya_tipini_bul(yol: str | Path) -> str:
    """Dosyanin hangi kaynak ailesine ait oldugunu belirler.

    Donen deger modeller.KAYNAK_TIPLERI kumesindendir. Dosya acilamiyorsa
    veya hicbir ipucu eslesmiyorsa 'genel' doner (istisna firlatmaz).
    """
    p = Path(yol)
    try:
        metinler, blob = _ipuclarini_topla(p)
    except Exception:
        return "genel"

    # 1) Antik ham cari hareket dokumu
    if _iceriyor(blob, "cari hareket"):
        return "antik_cari"

    # 2) Yuzyil elle dagitilmis (santiye + ucus kolonlari birlikte)
    if _iceriyor(blob, "santiyesi") and _iceriyor(blob, "ucus guzergahi"):
        return "yuzyil_dagitilmis"
    if _iceriyor(blob, "ucus guzergahi") and _iceriyor(blob, "ucus bilgisi"):
        return "yuzyil_dagitilmis"

    # 3) Energo assessment
    if "katilimci" in metinler and "paket" in metinler:
        return "energo_assessment"

    # 4) Energo arabuluculuk
    if "arabulucu" in metinler or _iceriyor(blob, "arabulucu"):
        return "energo_arabulucu"

    # 5) Saglik kontrol listesi
    if _iceriyor(blob, "bordrolu liste") or _iceriyor(blob, "bordrosuz liste"):
        return "energo_saglik"
    if _iceriyor(blob, "tckn") and _iceriyor(blob, "saglik kontrol"):
        return "energo_saglik"

    # 6) Koc Universitesi katilimci listesi
    if "id" in metinler and _iceriyor(blob, "alt fonksiyon"):
        return "koc_katilimci"

    return "genel"


def oku_tip(yol: str | Path, tip: str) -> list[GiderSatiri]:
    """Verilen kaynak tipinin parser'ini calistirir.

    Raises:
        KeyError: tip taninmiyorsa.
    """
    return PARSERLAR[tip](yol)


def oku(yol: str | Path) -> list[GiderSatiri]:
    """Dosya tipini bulur ve dogru parser'i calistirir.

    Ozel parser hic satir uretmezse (sablon beklenenden farkliysa) genel
    parser'a duser; boylece bilinmeyen bir surum sessizce bos sonuc vermez.
    """
    p = Path(yol)
    tip = dosya_tipini_bul(p)
    satirlar = oku_tip(p, tip)
    if not satirlar and tip != "genel":
        satirlar = genel_oku(p)
    return satirlar
