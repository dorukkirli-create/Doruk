"""1C personel listesi: grup sirketlerini kapsayan ikincil personel defteri.

Neden gerekli: ana veri (``2025_2026_giris_cikis.xlsx``) sadece RHI ve UST LUGA
tuzel kisilerini kapsar. Seyahat ve egitim faturalarinda Renservis,
Renstroydetal, RC, One Tower, Top Tower gibi grup sirketlerinin personeli de
gecer ve ana veride bulunamaz.

1C personel listesi (``1C Personnel List ...xlsx``) tum grup sirketlerini
kapsar. Olculen katki: 17.517 isimli kayittan 5.234'u ana veride YOKTUR.
Bunlarin dagilimi RSS 2.661, RC 1.550, UST LUGA 938, BSK 36, RHI 31, BSA 18.

Iki dosya AYNI ID uzayini kullanir (12.283 ortak sicil), bu yuzden birinde
bulunan bir kisi digerinde de ayni sicille aranabilir.

Onemli fark: ana veri AYLIK SNAPSHOT serisidir (donem bazli), 1C listesi ise
TEK bir tarihe ait durumdur. Bu yuzden 1C listesi ana verinin YERINE gecmez,
sadece ana veride bulunamayan kisiler icin basvurulur. Bu defterden gelen bir
kayit her zaman "donem dogrulanamadi" uyarisi tasir.

Kolon eslesmesi (Ingilizce basliklar, esnek cozulur)::

    ID              -> sicil
    Name Surname    -> ad_soyad
    Project         -> gorev_yeri      (masraf merkezi kaynagi)
    Firm            -> sirket
    Firm 2          -> sirket2
    Status          -> statu
    Collar          -> yaka
    Position        -> pozisyon
    Date of Birth   -> dogum_tarihi
    Work Status     -> calisma_durumu  (Rusca: Работает = calisiyor)
    Category        -> kategori_1c     (RHI / Other Companies / Excluded / ...)
"""

from __future__ import annotations

import logging
import pickle
from datetime import date, datetime
from pathlib import Path
from typing import Any

from masraf.metin import ascii_katla, isim_normalize, isim_tokenlari

_log = logging.getLogger(__name__)

ONBELLEK_SURUMU = 1

#: Beklenen alan -> kaynak dosyadaki kolon adinin ASCII/kucuk hali.
_KOLONLAR: dict[str, tuple[str, ...]] = {
    "sicil": ("id", "sicil", "sicil no"),
    "ad_soyad": ("name surname", "adi soyadi", "ad soyad", "name"),
    "pozisyon": ("position", "pozisyon"),
    "gorev_yeri": ("project", "proje"),
    "dep1": ("department", "departman"),
    "statu": ("status", "statu"),
    "yaka": ("collar", "yaka"),
    "sirket": ("firm", "sirket"),
    "sirket2": ("firm 2", "firm2", "sirket 2"),
    "vatandaslik": ("nationality", "vatandaslik"),
    "cinsiyet": ("gender", "cinsiyet"),
    "dogum_tarihi": ("date of birth", "dogum tarihi"),
    "ise_giris_tarihi": ("date of employment", "ise giris tarihi"),
    "sehir": ("location", "city", "sehir"),
    "calisma_durumu": ("work status", "calisma durumu"),
    "kategori_1c": ("category", "kategori"),
    "unvan": ("title", "unvan"),
}

#: Bu kategorideki satirlar gercek kisi degildir, indekse alinmaz.
_HARIC_KATEGORILER = frozenset({"excluded", "e data operator", "edataoperator"})


def _anahtar(ad: Any) -> str:
    metin = ascii_katla(str(ad)).replace("\n", " ").replace("_", " ")
    return " ".join(metin.lower().split())


def _sicil_normalize(deger: Any) -> str:
    if deger is None:
        return ""
    if isinstance(deger, float):
        if deger != deger:  # NaN
            return ""
        if deger.is_integer():
            return str(int(deger))
    metin = str(deger).strip()
    if not metin or metin.lower() in {"nan", "nat", "none"}:
        return ""
    if metin.endswith(".0") and metin[:-2].isdigit():
        metin = metin[:-2]
    return metin


def _tarihe_cevir(deger: Any) -> date | None:
    if isinstance(deger, datetime):
        return deger.date()
    if isinstance(deger, date):
        return deger
    return None


def _metin(deger: Any) -> str | None:
    if deger is None:
        return None
    try:
        if deger != deger:  # NaN / NaT
            return None
    except (TypeError, ValueError):
        pass
    metin = str(deger).strip()
    if not metin or metin.lower() in {"nan", "nat", "none"}:
        return None
    return metin


class YardimciDefter:
    """1C personel listesi uzerinde isim ve sicil aramasi sunar.

    Arayuzu ``PersonelDefteri`` ile bilincli olarak ayni tutulmustur; boylece
    eslestirici ikisini de ayni sekilde sorgular.
    """

    def __init__(self, kayitlar: list[dict[str, Any]], kaynak: str | None = None) -> None:
        self._kayitlar = kayitlar
        self.kaynak = kaynak
        self._sicil: dict[str, dict] = {}
        self._isim: dict[str, list[str]] = {}
        self._token: dict[frozenset[str], list[str]] = {}
        self._soyad: dict[str, list[str]] = {}
        self._indeksle()

    # ------------------------------------------------------------------

    def _indeksle(self) -> None:
        for kayit in self._kayitlar:
            sicil = kayit.get("sicil") or ""
            if not sicil:
                continue
            self._sicil.setdefault(sicil, kayit)

            ad = kayit.get("ad_soyad")
            if not ad:
                continue
            kategori = _anahtar(kayit.get("kategori_1c") or "")
            if kategori in _HARIC_KATEGORILER:
                continue

            norm = isim_normalize(ad)
            if not norm or norm == "NAN":
                continue
            kayit["ad_soyad_norm"] = norm
            self._isim.setdefault(norm, []).append(sicil)
            tokenlar = isim_tokenlari(ad)
            if tokenlar:
                self._token.setdefault(tokenlar, []).append(sicil)
                ilk = norm.split(" ")[0]
                if len(ilk) >= 3:
                    self._soyad.setdefault(ilk, []).append(sicil)

    # ------------------------------------------------------------------
    # Yukleme
    # ------------------------------------------------------------------

    @classmethod
    def yukle(cls, yol: str | Path, onbellek: bool = True) -> "YardimciDefter":
        """1C listesini okur. Onbellek varsa ve tazeyse ondan yukler."""
        hedef = Path(yol)
        if not hedef.is_file():
            raise FileNotFoundError(f"1C personel listesi bulunamadi: {hedef}")

        ob = hedef.with_suffix(hedef.suffix + ".yardimci.pkl")
        imza = (hedef.stat().st_mtime_ns, hedef.stat().st_size, ONBELLEK_SURUMU)
        if onbellek and ob.is_file():
            try:
                with open(ob, "rb") as f:
                    kayitli = pickle.load(f)
                if kayitli.get("imza") == imza:
                    return cls(kayitli["kayitlar"], str(hedef))
            except Exception as e:  # noqa: BLE001
                _log.warning("Yardimci defter onbellegi okunamadi: %s", e)

        import pandas as pd

        sayfa = 0
        try:
            sayfalar = pd.ExcelFile(hedef).sheet_names
            for aday in sayfalar:
                if _anahtar(aday) in {"main list", "ana liste", "liste"}:
                    sayfa = aday
                    break
        except Exception:  # noqa: BLE001
            pass

        df = pd.read_excel(hedef, sheet_name=sayfa, engine="openpyxl")
        cozum: dict[str, str] = {}
        mevcut = {_anahtar(k): k for k in df.columns}
        for alan, adaylar in _KOLONLAR.items():
            for aday in adaylar:
                if aday in mevcut:
                    cozum[alan] = mevcut[aday]
                    break
        eksik = [a for a in ("sicil", "ad_soyad", "gorev_yeri") if a not in cozum]
        if eksik:
            raise ValueError(
                "1C personel listesinde zorunlu kolonlar bulunamadi: " + ", ".join(eksik)
            )

        alt = df[[cozum[a] for a in cozum]].copy()
        alt.columns = list(cozum.keys())
        kayitlar: list[dict[str, Any]] = []
        for ham in alt.to_dict("records"):
            kayit: dict[str, Any] = {}
            for alan in cozum:
                deger = ham.get(alan)
                if alan == "sicil":
                    kayit[alan] = _sicil_normalize(deger)
                elif alan in ("dogum_tarihi", "ise_giris_tarihi"):
                    kayit[alan] = _tarihe_cevir(deger)
                else:
                    kayit[alan] = _metin(deger)
            kayit["_kaynak"] = "1c"
            kayitlar.append(kayit)

        if onbellek:
            try:
                with open(ob, "wb") as f:
                    pickle.dump({"imza": imza, "kayitlar": kayitlar}, f,
                                protocol=pickle.HIGHEST_PROTOCOL)
            except OSError as e:
                _log.warning("Yardimci defter onbellegi yazilamadi: %s", e)

        return cls(kayitlar, str(hedef))

    # ------------------------------------------------------------------
    # Sorgular (PersonelDefteri ile ayni imzalar)
    # ------------------------------------------------------------------

    def sicil_ile(self, sicil: str) -> dict | None:
        return self._sicil.get(_sicil_normalize(sicil))

    def isimle_adaylar(self, isim_norm: str) -> list[str]:
        return list(self._isim.get(isim_norm, ()))

    def token_ile_adaylar(self, tokenlar: frozenset) -> list[str]:
        return list(self._token.get(tokenlar, ()))

    def soyad_ile_adaylar(self, soyad: str) -> list[str]:
        return list(self._soyad.get(soyad, ()))

    def donem_kaydi(self, sicil: str, tarih: date | None) -> dict | None:
        """1C listesi tek bir ana ait oldugu icin donem secimi yapilmaz.

        Donen kayit her zaman ``_donem_eslesme = 'yardimci_defter'`` tasir;
        masraf merkezi cozumleyici bunu gorup uyari uretir.
        """
        kayit = self.sicil_ile(sicil)
        if kayit is None:
            return None
        kopya = dict(kayit)
        kopya["_donem_eslesme"] = "yardimci_defter"
        kopya["_donem_tahmini"] = True
        kopya.setdefault("donem", None)
        kopya.setdefault("kategori", None)
        kopya.setdefault("cikis_tarihi", None)
        return kopya

    @property
    def isim_sozlugu(self) -> set[str]:
        sozluk: set[str] = set()
        for tokenlar in self._token:
            sozluk.update(tokenlar)
        return sozluk

    def istatistik(self) -> dict:
        sirketler: dict[str, int] = {}
        for kayit in self._sicil.values():
            s = kayit.get("sirket2") or kayit.get("sirket") or "?"
            sirketler[s] = sirketler.get(s, 0) + 1
        return {
            "kayit_sayisi": len(self._kayitlar),
            "benzersiz_sicil": len(self._sicil),
            "isimli_sicil": sum(len(v) for v in self._isim.values()),
            "benzersiz_isim": len(self._isim),
            "sirket_dagilimi": dict(sorted(sirketler.items(), key=lambda p: -p[1])[:10]),
            "kaynak": self.kaynak,
        }
