"""Gercek veriye bagli test beklentileri ('altin ornekler').

Gercek personel adlari ve sicil numaralari DEPOYA GIRMEZ. Bunlara dayanan
testler beklentilerini ``ornek_veri/altin.json`` dosyasindan okur; dosya
yoksa test atlanir. Boylece test dosyalari kisisel veri tasimaz, gercek
veriye sahip makinede ise ayni testler tam guclerinde calisir.
"""

from __future__ import annotations

import json
import unittest
from pathlib import Path
from typing import Any

KOK = Path(__file__).resolve().parents[1]
ALTIN_YOLU = KOK / "ornek_veri" / "altin.json"

_ONBELLEK: dict[str, Any] = {}


def altin_veya_none(*anahtar: str) -> Any:
    """Altin dosyasindan degeri dondurur; dosya ya da anahtar yoksa None."""
    if "veri" not in _ONBELLEK:
        try:
            _ONBELLEK["veri"] = json.loads(ALTIN_YOLU.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            _ONBELLEK["veri"] = None
    deger = _ONBELLEK["veri"]
    for a in anahtar:
        if not isinstance(deger, dict) or a not in deger:
            return None
        deger = deger[a]
    return deger


def altin(*anahtar: str) -> Any:
    """Degeri dondurur; yoksa testi atlar (unittest.SkipTest)."""
    deger = altin_veya_none(*anahtar)
    if deger is None:
        raise unittest.SkipTest(
            f"altin ornek yok: {'/'.join(anahtar)} ({ALTIN_YOLU.name} bulunamadi)")
    return deger
