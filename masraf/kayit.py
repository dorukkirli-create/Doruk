"""Personel ana verisi defteri: yukleme, onbellek ve eslestirme indeksleri.

Kaynak dosya aylik SNAPSHOT yapisindadir: her kisi her 'Donem' icin bir satir
tasir. Bu modul dosyayi okur, kolon adlarini esnek cozer, sicil/isim/token/soyad
indekslerini kurar ve bir gider tarihine karsilik gelen donem kaydini bulur.

Onemli veri gercekleri (olculmus):
- 'Sicil' kolonu karisik tiptedir (str + int) ve 'C92620' gibi harfli sicil
  icerir. Her zaman str'e cevrilip strip edilir, '.0' eki temizlenir.
- 'Adi Soyadi' bos olan satirlar 'Bordrosuz Taseron' kayitlaridir; sahte sicil
  tasirlar ve isim indekslerine EKLENMEZ (sicil indeksine eklenir).
- 'Kategori' iki farkli yazimla gelir ('Cikis' ve Turkce karakterli hali);
  normalize edilir.
- Personel ana verisinde TCKN ve pasaport numarasi YOKTUR. TCKN eslestirmesi
  icin ayri bir kopru tablosu (veri/tckn_sicil.csv) gerekir.
"""

from __future__ import annotations

import bisect
import pickle
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from masraf.metin import ascii_katla, isim_normalize, isim_tokenlari

# Onbellek bicim surumu. Indeks yapisi degisirse artirilmalidir; eski onbellek
# dosyalari otomatik olarak gecersiz sayilir.
ONBELLEK_SURUMU = 4

# Beklenen alan adi -> kaynak dosyadaki kolon adinin ASCII/kucuk hali.
# Kolon adlari dosya versiyonuna gore degisebilecegi icin esnek cozulur.
_KOLON_HARITASI: dict[str, tuple[str, ...]] = {
    "sicil": ("sicil",),
    "ad_soyad": ("adi soyadi", "ad soyad", "adisoyadi"),
    "pozisyon": ("pozisyon",),
    "gorev_yeri": ("gorev yeri", "gorevyeri"),
    "dep1": ("dep1",),
    "statu": ("statu",),
    "yaka": ("yaka",),
    "sirket": ("sirket",),
    "vatandaslik": ("vatandaslik",),
    "cinsiyet": ("cinsiyet",),
    "dogum_tarihi": ("dogum tarihi",),
    "ise_giris_tarihi": ("rhi ise giris tarihi", "ise giris tarihi"),
    "sehir": ("sehir",),
    "sirket2": ("sirket 2", "sirket2"),
    "kademe": ("kademe",),
    "donem": ("donem",),
    "kategori": ("kategori",),
    "aktif": ("aktif",),
    "ayrilan": ("ayrilan",),
    "istifa_ile_ayrilan": ("istifa ile ayrilan",),
    "rhi_taseron": ("rhi taseron",),
    "cikis_tarihi": ("isten cikis tarihi",),
    "cikis_sebebi_resmi": ("isten cikis sebebi resmi",),
    "cikis_sebebi_reel": ("isten cikis sebebi reel",),
    "cikis_ik_not": ("isten cikis ik not",),
    "cikis_yonetici_not": ("isten cikis yonetici not",),
    "cikis_yonetici_kim": ("isten cikis yonetici kim",),
    "yas": ("yas",),
    "kidem": ("kidem",),
    "is_ailesi": ("ek1 is ailesi disiplin", "ek1 is ailesi/disiplin", "is ailesi"),
    "usta": ("ek2 usta ustabasi", "ek2 usta/ustabasi"),
}

# Tarih olarak yorumlanacak alanlar.
_TARIH_ALANLARI = ("dogum_tarihi", "ise_giris_tarihi", "donem", "cikis_tarihi")

# Bordrosuz taseron isareti (ASCII katlanmis, kucuk harf karsilastirmasi).
_BORDROSUZ = "bordrosuz taseron"


def _kolon_anahtari(ad: Any) -> str:
    """Kolon adini karsilastirilabilir ASCII/kucuk bicime cevirir."""
    metin = ascii_katla(str(ad))
    metin = metin.replace("\n", " ").replace("\r", " ").replace("_", " ")
    metin = metin.replace("/", " ").replace("-", " ")
    return " ".join(metin.lower().split())


def sicil_normalize(deger: Any) -> str:
    """Sicil degerini kanonik str bicime cevirir.

    'Sicil' kolonu karisik tiptedir; int okunan degerler pandas tarafindan
    632481.0 gibi float'a donusebilir. Bu fonksiyon her zaman kullanilmalidir,
    aksi halde join'ler SESSIZCE bozulur.
    """
    if deger is None:
        return ""
    if isinstance(deger, float):
        if pd.isna(deger):
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
    """Cesitli tarih temsillerini date'e cevirir; cozulemezse None."""
    if deger is None:
        return None
    if isinstance(deger, datetime):
        return deger.date()
    if isinstance(deger, date):
        return deger
    try:
        if pd.isna(deger):
            return None
    except (TypeError, ValueError):
        pass
    try:
        zaman = pd.Timestamp(deger)
    except (TypeError, ValueError):
        return None
    if pd.isna(zaman):
        return None
    return zaman.date()


def _metin_temizle(deger: Any) -> str | None:
    """Hucre degerini duz metne cevirir; bos ise None."""
    if deger is None:
        return None
    try:
        if pd.isna(deger):
            return None
    except (TypeError, ValueError):
        pass
    metin = str(deger).strip()
    if not metin or metin.lower() in {"nan", "nat", "none"}:
        return None
    return metin


def _kategori_normalize(deger: Any) -> str | None:
    """'Cikis' / Turkce karakterli 'Cikis' yazimlarini tek bicime indirger."""
    metin = _metin_temizle(deger)
    if metin is None:
        return None
    katlanmis = ascii_katla(metin).strip().lower()
    if katlanmis.startswith("cikis"):
        return "Cikis"
    if katlanmis.startswith("aktif"):
        return "Aktif"
    return metin


class PersonelDefteri:
    """Personel ana verisi uzerinde eslestirme indeksleri sunan defter.

    Tum indeksler __init__ icinde kurulur. Kayitlar sicil basina donem sirali
    tutulur, boylece bir gider tarihine karsilik gelen donem bisect ile
    bulunabilir.
    """

    def __init__(self, df: pd.DataFrame) -> None:
        self._kayitlar: list[dict[str, Any]] = []
        self._sicil_konumlar: dict[str, list[int]] = {}
        self._sicil_donemler: dict[str, list[date]] = {}
        self._isim_index: dict[str, list[str]] = {}
        self._token_index: dict[frozenset[str], list[str]] = {}
        self._soyad_index: dict[str, list[str]] = {}
        self._isim_sozlugu: set[str] = set()
        self._kaynak_yol: str | None = None
        self._insa_et(df)

    # ------------------------------------------------------------------
    # Insa
    # ------------------------------------------------------------------

    def _insa_et(self, df: pd.DataFrame) -> None:
        cozum = self._kolonlari_coz(df)
        eksik = [ad for ad in ("sicil", "ad_soyad", "gorev_yeri", "donem") if ad not in cozum]
        if eksik:
            raise ValueError(
                f"Personel dosyasinda zorunlu kolonlar bulunamadi: {', '.join(eksik)}"
            )

        # Sadece cozulmus kolonlari al, beklenen adlara yeniden adlandir.
        alt = df[[cozum[ad] for ad in cozum]].copy()
        alt.columns = list(cozum.keys())

        kayitlar = alt.to_dict("records")
        for ham in kayitlar:
            kayit: dict[str, Any] = {}
            for alan in cozum:
                deger = ham.get(alan)
                if alan in _TARIH_ALANLARI:
                    kayit[alan] = _tarihe_cevir(deger)
                elif alan == "sicil":
                    kayit[alan] = sicil_normalize(deger)
                elif alan == "kategori":
                    kayit[alan] = _kategori_normalize(deger)
                elif alan in ("kademe", "aktif", "ayrilan", "istifa_ile_ayrilan", "yas", "kidem"):
                    kayit[alan] = deger if not pd.isna(deger) else None
                else:
                    kayit[alan] = _metin_temizle(deger)

            sicil = kayit["sicil"]
            if not sicil:
                continue

            ad_soyad = kayit.get("ad_soyad")
            taseron = kayit.get("rhi_taseron")
            bordrosuz = bool(taseron and ascii_katla(taseron).strip().lower() == _BORDROSUZ)
            kayit["bordrosuz"] = bordrosuz
            kayit["ad_soyad_norm"] = isim_normalize(ad_soyad) if ad_soyad else ""

            konum = len(self._kayitlar)
            self._kayitlar.append(kayit)
            self._sicil_konumlar.setdefault(sicil, []).append(konum)

        # Sicil basina donem sirasina diz ve paralel donem listesi kur.
        for sicil, konumlar in self._sicil_konumlar.items():
            konumlar.sort(
                key=lambda k: (self._kayitlar[k]["donem"] is None, self._kayitlar[k]["donem"])
            )
            self._sicil_donemler[sicil] = [
                self._kayitlar[k]["donem"] for k in konumlar if self._kayitlar[k]["donem"] is not None
            ]

        self._isim_indekslerini_kur()

    def _isim_indekslerini_kur(self) -> None:
        """Isim / token / soyad indekslerini kurar.

        Adi bos olan satirlar (bordrosuz taseron kayitlari) isim indekslerine
        EKLENMEZ; sicil indeksinde ise zaten yer alirlar.
        Personel ana verisinde isimler 'SOYAD AD' sirasindadir, bu nedenle
        soyad indeksi ilk token uzerinden kurulur.

        DIKKAT: Bir kisinin adi donemler arasinda DEGISEBILIR (orn. evlilik
        sonrasi soyadi degisikligi: sicil 625770 'Coskun Ese' -> 'Yaprak Ese').
        Indeksler kisinin TUM donemlerdeki yazimlarini kapsar, bu nedenle eski
        adiyla kesilmis bir fatura da dogru sicile eslesir. sicil_ile() ise her
        zaman EN GUNCEL adi dondurur; bu yuzden eslesen isim ile dondurulen ad
        farkli olabilir, bu bir hata degildir.
        """
        isim_gorulen: dict[str, set[str]] = {}
        token_gorulen: dict[frozenset[str], set[str]] = {}
        soyad_gorulen: dict[str, set[str]] = {}

        for kayit in self._kayitlar:
            norm = kayit["ad_soyad_norm"]
            if not norm:
                continue
            sicil = kayit["sicil"]
            tokenlar = frozenset(norm.split(" "))
            self._isim_sozlugu.update(tokenlar)

            if norm not in isim_gorulen:
                isim_gorulen[norm] = set()
                self._isim_index[norm] = []
            if sicil not in isim_gorulen[norm]:
                isim_gorulen[norm].add(sicil)
                self._isim_index[norm].append(sicil)

            if tokenlar not in token_gorulen:
                token_gorulen[tokenlar] = set()
                self._token_index[tokenlar] = []
            if sicil not in token_gorulen[tokenlar]:
                token_gorulen[tokenlar].add(sicil)
                self._token_index[tokenlar].append(sicil)

            soyad = norm.split(" ", 1)[0]
            if soyad not in soyad_gorulen:
                soyad_gorulen[soyad] = set()
                self._soyad_index[soyad] = []
            if sicil not in soyad_gorulen[soyad]:
                soyad_gorulen[soyad].add(sicil)
                self._soyad_index[soyad].append(sicil)

    @staticmethod
    def _kolonlari_coz(df: pd.DataFrame) -> dict[str, str]:
        """Gercek kolon adlarini beklenen alan adlarina esler."""
        mevcut = {_kolon_anahtari(c): c for c in df.columns}
        cozum: dict[str, str] = {}
        for alan, adaylar in _KOLON_HARITASI.items():
            for aday in adaylar:
                if aday in mevcut:
                    cozum[alan] = mevcut[aday]
                    break
        return cozum

    # ------------------------------------------------------------------
    # Yukleme ve onbellek
    # ------------------------------------------------------------------

    @classmethod
    def yukle(cls, yol: str | Path, onbellek: bool = True) -> "PersonelDefteri":
        """Personel ana verisini okur, mumkunse pickle onbelleginden yukler.

        24 MB'lik xlsx dosyasinin okunmasi ~25 sn surer. Onbellek dosyasi
        kaynagin yanina '<dosya>.onbellek.pkl' olarak yazilir; kaynak dosyanin
        mtime ve boyutu degistiyse onbellek gecersiz sayilir ve yeniden okunur.
        """
        kaynak = Path(yol)
        if not kaynak.exists():
            raise FileNotFoundError(f"Personel dosyasi bulunamadi: {kaynak}")

        bilgi = kaynak.stat()
        imza = (ONBELLEK_SURUMU, int(bilgi.st_mtime), bilgi.st_size)
        onbellek_yolu = kaynak.with_suffix(kaynak.suffix + ".onbellek.pkl")

        if onbellek and onbellek_yolu.exists():
            try:
                with onbellek_yolu.open("rb") as dosya:
                    paket = pickle.load(dosya)
                if paket.get("imza") == imza:
                    defter = paket["defter"]
                    defter._kaynak_yol = str(kaynak)
                    return defter
            except (pickle.PickleError, EOFError, KeyError, AttributeError, OSError):
                pass  # Bozuk/eski onbellek: sessizce yeniden oku.

        df = pd.read_excel(kaynak, sheet_name=0)
        defter = cls(df)
        defter._kaynak_yol = str(kaynak)

        if onbellek:
            try:
                gecici = onbellek_yolu.with_suffix(".pkl.tmp")
                with gecici.open("wb") as dosya:
                    pickle.dump(
                        {"imza": imza, "defter": defter}, dosya, protocol=pickle.HIGHEST_PROTOCOL
                    )
                gecici.replace(onbellek_yolu)
            except OSError:
                pass  # Onbellek yazilamazsa calismaya devam et.

        return defter

    # ------------------------------------------------------------------
    # Sorgular
    # ------------------------------------------------------------------

    def sicil_ile(self, sicil: str) -> dict | None:
        """Sicile ait EN GUNCEL donem kaydini dondurur; yoksa None."""
        anahtar = sicil_normalize(sicil)
        konumlar = self._sicil_konumlar.get(anahtar)
        if not konumlar:
            return None
        return dict(self._kayitlar[konumlar[-1]])

    def isimle_adaylar(self, isim_norm: str) -> list[str]:
        """Normalize edilmis isme birebir uyan sicil listesini dondurur."""
        if not isim_norm:
            return []
        return list(self._isim_index.get(isim_norm, ()))

    def token_ile_adaylar(self, tokenlar: frozenset) -> list[str]:
        """Token kumesi birebir ayni olan sicilleri dondurur.

        Sirasiz karsilastirma sayesinde 'SOYAD AD' ile 'AD SOYAD' yazimlari
        ayni adaylara ulasir.
        """
        if not tokenlar:
            return []
        return list(self._token_index.get(frozenset(tokenlar), ()))

    def soyad_ile_adaylar(self, soyad: str) -> list[str]:
        """Soyadi eslesen sicilleri dondurur (aile bireyi tespiti icin).

        Personel verisinde isimler 'SOYAD AD' sirasinda oldugundan indeks ilk
        token uzerinden kurulmustur.
        """
        anahtar = isim_normalize(soyad)
        if not anahtar:
            return []
        return list(self._soyad_index.get(anahtar.split(" ", 1)[0], ()))

    def donem_kaydi(self, sicil: str, tarih: date | None) -> dict | None:
        """Verilen tarihe ait donem kaydini dondurur.

        Tarihe esit veya ondan kucuk EN BUYUK donem secilir. Tarih tum
        donemlerden onceyse EN ERKEN donem dondurulur ve kayda
        {'_donem_tahmini': True} isareti konur. Tarih None ise en guncel
        kayit dondurulur ve yine tahmin olarak isaretlenir.
        """
        anahtar = sicil_normalize(sicil)
        konumlar = self._sicil_konumlar.get(anahtar)
        if not konumlar:
            return None

        donemler = self._sicil_donemler.get(anahtar, [])
        if tarih is None or not donemler:
            kayit = dict(self._kayitlar[konumlar[-1]])
            kayit["_donem_tahmini"] = True
            return kayit

        hedef = _tarihe_cevir(tarih)
        if hedef is None:
            kayit = dict(self._kayitlar[konumlar[-1]])
            kayit["_donem_tahmini"] = True
            return kayit

        yer = bisect.bisect_right(donemler, hedef)
        if yer == 0:
            # Tarih tum donemlerden once: en erken donemi dondur, tahmin isaretle.
            kayit = dict(self._kayitlar[konumlar[0]])
            kayit["_donem_tahmini"] = True
            return kayit

        kayit = dict(self._kayitlar[konumlar[yer - 1]])
        kayit["_donem_tahmini"] = False
        return kayit

    # ------------------------------------------------------------------
    # Yardimcilar
    # ------------------------------------------------------------------

    @property
    def isim_sozlugu(self) -> set[str]:
        """Tum isim tokenlarinin kumesi (bitisik ad acmak icin kullanilir)."""
        return self._isim_sozlugu

    @property
    def donemler(self) -> list[date]:
        """Veride gecen tum donemler, sirali."""
        gorulen: set[date] = set()
        for kayit in self._kayitlar:
            donem = kayit.get("donem")
            if donem is not None:
                gorulen.add(donem)
        return sorted(gorulen)

    @property
    def gorev_yerleri(self) -> list[str]:
        """Veride gecen tum 'Gorev Yeri' degerleri (masraf merkezi adaylari)."""
        gorulen: set[str] = set()
        for kayit in self._kayitlar:
            deger = kayit.get("gorev_yeri")
            if deger:
                gorulen.add(deger)
        return sorted(gorulen)

    def istatistik(self) -> dict:
        """Defterin ozet istatistiklerini dondurur (dogrulama ve arayuz icin)."""
        donemler = self.donemler
        isimli_siciller = {
            k["sicil"] for k in self._kayitlar if k["ad_soyad_norm"]
        }
        isimsiz_satir = sum(1 for k in self._kayitlar if not k["ad_soyad_norm"])
        cakisan = sum(1 for siciller in self._isim_index.values() if len(siciller) > 1)
        # Cakisan isimlerin kapsadigi sicil sayisi: bir gider satirinda bu
        # sicillerden birine denk gelme olasiligini gosterir.
        cakisan_sicil = sum(
            len(siciller) for siciller in self._isim_index.values() if len(siciller) > 1
        )
        return {
            "satir_sayisi": len(self._kayitlar),
            "benzersiz_sicil": len(self._sicil_konumlar),
            "isimli_sicil": len(isimli_siciller),
            "isimsiz_satir": isimsiz_satir,
            "benzersiz_isim": len(self._isim_index),
            "cakisan_isim": cakisan,
            "cakisma_orani": round(cakisan / len(self._isim_index) * 100, 2) if self._isim_index else 0.0,
            "cakisan_sicil": cakisan_sicil,
            "cakisan_sicil_orani": round(cakisan_sicil / len(isimli_siciller) * 100, 2) if isimli_siciller else 0.0,
            "donem_sayisi": len(donemler),
            "ilk_donem": donemler[0] if donemler else None,
            "son_donem": donemler[-1] if donemler else None,
            "gorev_yeri_sayisi": len(self.gorev_yerleri),
            "isim_sozlugu_boyutu": len(self._isim_sozlugu),
            "kaynak": self._kaynak_yol,
        }
