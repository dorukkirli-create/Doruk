"""Kaynak fatura dosyasi okuyuculari (parser'lar).

Her parser bir dosya yolunu alir ve modeller.GiderSatiri listesi dondurur.
Moduller katmanlidir ve dairesel bagimlilik icermez:

    genel   -> yardimci hucre/kolon cozumleyicileri + taninmayan dosya parser'i
    antik   -> Antik / Yuzyil Travel dosyalari      (genel'i kullanir)
    energo  -> Energo yansitma dosyalari            (genel'i kullanir)
    kesif   -> dosya tipi tespiti ve yonlendirme    (hepsini kullanir)

Tipik kullanim:

    from masraf.okuyucular.kesif import oku
    satirlar = oku("gelen/ANTIK_CARI_TEMMUZ_2026.xls")

Paket duzeyindeki isimler TEMBEL yuklenir; 'from masraf.okuyucular import
antik' gibi dogrudan modul importlari da desteklenir.
"""

from __future__ import annotations

from typing import Any

__all__ = [
    "antik_cari_oku",
    "yuzyil_dagitilmis_oku",
    "assessment_oku",
    "arabulucu_oku",
    "saglik_oku",
    "koc_katilimci_oku",
    "genel_oku",
    "dosya_tipini_bul",
    "oku",
    "oku_tip",
]

# Ad -> (modul adi) esleme; tembel yukleme icin.
_KAYNAKLAR: dict[str, str] = {
    "antik_cari_oku": "antik",
    "yuzyil_dagitilmis_oku": "antik",
    "assessment_oku": "energo",
    "arabulucu_oku": "energo",
    "saglik_oku": "energo",
    "koc_katilimci_oku": "energo",
    "genel_oku": "genel",
    "dosya_tipini_bul": "kesif",
    "oku": "kesif",
    "oku_tip": "kesif",
}


def __getattr__(ad: str) -> Any:
    """Paket duzeyindeki isimleri ilk erisimde ilgili modulden yukler."""
    modul_adi = _KAYNAKLAR.get(ad)
    if modul_adi is None:
        raise AttributeError(f"module {__name__!r} has no attribute {ad!r}")
    from importlib import import_module

    modul = import_module(f"{__name__}.{modul_adi}")
    deger = getattr(modul, ad)
    globals()[ad] = deger
    return deger


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))
