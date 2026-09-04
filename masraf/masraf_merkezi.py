"""Masraf merkezi cozumleme: gorev yeri -> finans masraf merkezi kodu.

Bu modul iki isi yapar:

1. ``MasrafMerkeziHaritasi`` — ``veri/masraf_merkezi_haritasi.csv`` dosyasini
   okur ve personel ana verisindeki 'Gorev Yeri' degerlerini finans tarafindaki
   masraf merkezi kodlarina cevirir. Kullanici bu CSV'yi kendi kodlariyla
   degistirebilir; kod hicbir kodu sabit olarak icermez.

2. ``masraf_merkezi_coz`` — bir gider satirini, onun eslesmesini ve personel
   defterini alip nihai ``Sonuc`` kaydini uretir. Buradaki en kritik nokta
   TARIHE GORE DONEM SECIMIdir: personel dosyasi aylik snapshot'tir ve
   kisilerin yaklasik yuzde 1,4'unun gorev yeri donemler arasinda degisir.
   Belge tarihindeki donem kaydi kullanilmazsa tam o kisiler yanlis masraf
   merkezine mahsuplasir.

Modul tek basina import edilebilir; pandas veya eslestirici modullerine
bagimli degildir.

Kaynak dosyadaki santiye etiketi hakkinda
-----------------------------------------
Bazi kaynak dosyalarda insan tarafindan doldurulmus bir santiye/proje kolonu
bulunur (``GiderSatiri.masraf_merkezi_kaynak``). Bu etiket iki farkli seyi
karistirarak icerir:

* PROJE adi ('GPP Proje', 'Udokan GMK', 'AMURSKI GAZ ISLETME FABRIKASI')
* TUZEL KISI adi ('RHI', 'RENSERVIS', 'ONE TOWER', 'RC PETER')

Proje etiketi ile hesaplanan masraf merkezi celisiyorsa bu gercek bir bulgudur
ve satir incelemeye gonderilir (elle yapilan hatalari yakalar). Tuzel kisi
etiketi ise proje ile karsilastirilamaz; onun karsiligi personel kaydindaki
'Sirket 2' kolonudur ve varsayilan olarak UYARI URETMEZ, cunku elle etiketleme
kendi icinde tutarsizdir (ayni projede calisan kisiler bir yerde 'RHI', baska
yerde 'UST LUGA GPP' yazilmistir). Bu karsilastirma
``tuzel_kisi_uyar=True`` ile acilabilir.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from masraf.metin import ascii_katla
from masraf.modeller import (
    DURUM_ESLESMEDI,
    DURUM_INCELE,
    DURUM_OTOMATIK,
    Eslesme,
    GiderSatiri,
    Sonuc,
)

__all__ = [
    "MasrafMerkeziHaritasi",
    "masraf_merkezi_coz",
    "varsayilan_harita_yolu",
    "TUZEL_KISI_ETIKETLERI",
    "EK_ESANLAMLILAR",
    "GUVEN_ESIGI",
    "ALT_ESIK",
]

#: Bu esigin ustundeki ve uyarisiz satirlar otomatik kabul edilir.
GUVEN_ESIGI = 0.90

#: Bu esigin altindaki satirlar eslesmemis sayilir.
ALT_ESIK = 0.50

#: Beklenen CSV kolonlari.
BASLIKLAR: tuple[str, ...] = (
    "gorev_yeri",
    "masraf_merkezi_kodu",
    "masraf_merkezi_adi",
    "sirket",
    "aktif",
)

#: Varsayilan harita dosyasinin adi.
DOSYA_ADI = "masraf_merkezi_haritasi.csv"

#: Kaynak dosyalardaki santiye kolonunda gecen TUZEL KISI etiketleri.
#: Bunlar proje degildir; masraf merkezi ile karsilastirilmazlar.
TUZEL_KISI_ETIKETLERI: frozenset[str] = frozenset({
    "RHI",
    "RHI RUSSIA",
    "UST LUGA",
    "USTLUGA",
    "RENSERVIS",
    "RENSTROYDETAL",
    "RSD",
    "RS",
    "RC PETER",
    "RC PETERSBURG",
    "RC MOSKOVA",
    "RC MOSCOW",
    "ONE TOWER",
    "TOP TOWER",
    "SAREN",
    "KRONDEX",
    "YAKA LLC",
    "YAKA",
})

#: Kaynak dosyalarda gecen proje yazimlarinin personel verisindeki 'Gorev Yeri'
#: karsiliklari. Kullanici CSV'sini kirletmemek icin kod tarafinda tutulur;
#: CSV'ye ayni gorev yeri icin ikinci bir satir eklenmesi de calisir.
EK_ESANLAMLILAR: dict[str, str] = {
    "GPP PROJE": "GPP Project",
    "GPP PROJESI": "GPP Project",
    "UST LUGA GPP": "GPP Project",
    "UST LUGA GPP PROJESI": "GPP Project",
    "UST LUGA GAS PROCESSING COMPLEX GPP": "GPP Project",
    "UST LUGA GAS PROCESSING COMPLEX GPC": "Ust Luga Fabrication - GPC (RHI)",
    "AMURSKI GAZ ISLETME FABRIKASI": "Amursky Gas Processing Plant",
    "AMURSKIY GAZ ISLETME FABRIKASI": "Amursky Gas Processing Plant",
    "CATERING AMURSKY GAS PROCESSING PLANT": "Amursky Gas Processing Plant",
    "AMUR": "Amursky Gas Processing Plant",
    "AMURSKY": "Amursky Gas Processing Plant",
    "AMUR GPZ": "Amursky Gas Processing Plant",
    "AMUR AGPP": "Amursky Gas Processing Plant",
    "AGPZ": "Amursky Gas Processing Plant",
    "UDOKAN GMK": "Udokan (GMK)",
    "UDOKAN": "Udokan (GMK)",
    "GYDAN": "ALNG2-Gydan",
    "ALNG2 GYDAN": "ALNG2-Gydan",
    "ALNG2 GBS": "ALNG2-GBS Project",
    "MOSKOVA": "RHI Russia - Headquarter (Moscow)",
    "MOSCOW": "RHI Russia - Headquarter (Moscow)",
    "MOSKOVA OFIS": "RHI Russia - Headquarter (Moscow)",
    "HEADQUARTER": "RHI Russia - Headquarter (Moscow)",
    "UST LUGA ST PETERSBURG OFFICE": "Ust-Luga – Reshetnikova Office",
    "ST PETERSBURG OFFICE": "Ust-Luga – Reshetnikova Office",
    "RESHETNIKOVA": "Ust-Luga – Reshetnikova Office",
}

#: Tire benzeri karakterler (en-dash, em-dash, tire, alt cizgi, slash).
_TIRE_KARAKTERLERI = "-‐‑‒–—―_/\\"

#: 'aktif' kolonunda dogru sayilan degerler.
_DOGRU_DEGERLER = frozenset({"E", "EVET", "1", "TRUE", "T", "YES", "Y", "AKTIF", "X"})


def varsayilan_harita_yolu(veri_dizini: str | Path = "veri") -> Path:
    """Varsayilan harita dosyasinin yolunu dondurur."""
    return Path(veri_dizini) / DOSYA_ADI


def _anahtar(deger: Any) -> str:
    """Gorev yeri / masraf merkezi metnini karsilastirilabilir bicime cevirir.

    'Ust-Luga – Reshetnikova Office' ve 'Ust Luga - Reshetnikova Office'
    ayni anahtara duser: 'UST LUGA RESHETNIKOVA OFFICE'. Parantez, tire ve
    coklu bosluk farklari yok sayilir.
    """
    if deger is None:
        return ""
    metin = ascii_katla(str(deger)).upper()
    for karakter in _TIRE_KARAKTERLERI:
        metin = metin.replace(karakter, " ")
    metin = metin.replace("(", " ").replace(")", " ")
    metin = metin.replace(".", " ").replace(",", " ").replace("&", " ")
    return " ".join(metin.split())


def _metin(deger: Any) -> str:
    """Hucre degerini kirpilmis metne cevirir."""
    if deger is None:
        return ""
    metin = str(deger).strip()
    if metin.lower() in {"nan", "nat", "none"}:
        return ""
    return metin


def _dogru_mu(deger: Any, varsayilan: bool = True) -> bool:
    """'aktif' kolonunu yorumlar; bos ise varsayilani dondurur."""
    metin = _anahtar(deger)
    if not metin:
        return varsayilan
    if metin in _DOGRU_DEGERLER:
        return True
    return False


def _gecerli_tarih(deger: Any) -> date | None:
    """Tarih benzeri degeri gecerli bir date'e indirger; degilse None.

    DIKKAT: pandas NaT degeri ``datetime`` alt sinifidir; ``is not None``
    kontrolunu GECER ama .year erisimi hata verir ve karsilastirmalari her
    zaman False doner. Bu yuzden 'kendine esit degil' testi ile elenir
    (NaT != NaT dogrudur). Boylece pandas'a bagimlilik da olusmaz.
    """
    if deger is None:
        return None
    try:
        if deger != deger:  # NaT / NaN
            return None
    except (TypeError, ValueError):
        return None
    if isinstance(deger, datetime):
        return deger.date()
    if isinstance(deger, date):
        return deger
    return None


def _tarih_metni(deger: date | None) -> str:
    """Tarihi GG.AA.YYYY olarak yazar; None ise bos metin."""
    gecerli = _gecerli_tarih(deger)
    if gecerli is None:
        return ""
    return gecerli.strftime("%d.%m.%Y")


def _ay_basi_sonrasi(tarih: date, donem: date) -> bool:
    """Belge tarihi verilen donemin ayindan SONRAKI bir aya mi dusuyor?"""
    return (tarih.year, tarih.month) > (donem.year, donem.month)


@dataclass(frozen=True)
class MasrafMerkezi:
    """Harita dosyasindaki tek bir satir."""

    gorev_yeri: str
    kod: str
    ad: str
    sirket: str | None
    aktif: bool

    def sozluk(self) -> dict:
        """coz() tarafindan dondurulen sozluk bicimi."""
        return {
            "gorev_yeri": self.gorev_yeri,
            "masraf_merkezi_kodu": self.kod,
            "masraf_merkezi_adi": self.ad,
            "sirket": self.sirket,
            "aktif": self.aktif,
        }


class MasrafMerkeziHaritasi:
    """Gorev yeri -> masraf merkezi kodu esleme tablosu.

    Arama sirasi:
        1. Gorev yerinin birebir kendisi
        2. Normalize anahtar (tire/parantez/bosluk farklari yok sayilir)
        3. Masraf merkezi kodunun kendisi (kullanici kodu yazmis olabilir)
        4. Masraf merkezi adi
        5. Kod icindeki esanlamli tablosu (kaynak dosyalardaki proje yazimlari)

    Dosya yoksa BOS harita dondurulur ve ``kaynak_var`` False olur; boru hatti
    bu durumda gorev yerini oldugu gibi kullanip uyari yazar, calismaya devam
    eder.
    """

    def __init__(self, kayitlar: Iterable[MasrafMerkezi] | None = None,
                 kaynak: str | None = None, kaynak_var: bool = True) -> None:
        self.kaynak = kaynak
        self.kaynak_var = kaynak_var
        self._kayitlar: list[MasrafMerkezi] = list(kayitlar or ())
        self._index: dict[str, MasrafMerkezi] = {}
        self._bilinmeyenler: set[str] = set()
        self._indeksle()

    # ------------------------------------------------------------------
    # Kurulum
    # ------------------------------------------------------------------

    def _indeksle(self) -> None:
        """Arama indeksini kurar. Once yazilan kayit onceliklidir."""
        self._index.clear()
        for kayit in self._kayitlar:
            for aday in (kayit.gorev_yeri, kayit.kod, kayit.ad):
                anahtar = _anahtar(aday)
                if anahtar:
                    self._index.setdefault(anahtar, kayit)
        # Esanlamlilar: hedef gorev yeri haritada varsa onun kaydina baglanir.
        for yazim, hedef in EK_ESANLAMLILAR.items():
            kayit = self._index.get(_anahtar(hedef))
            if kayit is not None:
                self._index.setdefault(_anahtar(yazim), kayit)

    @classmethod
    def yukle(cls, yol: str | Path) -> "MasrafMerkeziHaritasi":
        """CSV harita dosyasini okur.

        Ayirici otomatik bulunur (virgul veya noktali virgul); Excel'in
        Turkce/Rusca yerel ayarla kaydettigi dosyalar da okunur.
        """
        hedef = Path(yol)
        if not hedef.exists():
            return cls([], kaynak=str(hedef), kaynak_var=False)

        try:
            ham = hedef.read_text(encoding="utf-8-sig")
        except (OSError, UnicodeDecodeError):
            try:
                ham = hedef.read_text(encoding="utf-8", errors="replace")
            except OSError:
                return cls([], kaynak=str(hedef), kaynak_var=False)

        satirlar = [s for s in ham.splitlines() if s.strip()]
        if not satirlar:
            return cls([], kaynak=str(hedef), kaynak_var=False)

        ilk = satirlar[0]
        ayirici = ";" if ilk.count(";") > ilk.count(",") else ","
        okuyucu = csv.DictReader(satirlar, delimiter=ayirici)

        # Kolon adlarini esnek coz (bosluk/buyuk harf farkina takilma).
        alan_haritasi: dict[str, str] = {}
        for alan in okuyucu.fieldnames or []:
            anahtar = _anahtar(alan).replace(" ", "_").lower()
            alan_haritasi[anahtar] = alan

        def al(satir: dict, ad: str) -> str:
            gercek = alan_haritasi.get(ad)
            return _metin(satir.get(gercek)) if gercek else ""

        kayitlar: list[MasrafMerkezi] = []
        for satir in okuyucu:
            gorev_yeri = al(satir, "gorev_yeri")
            if not gorev_yeri:
                continue
            kod = al(satir, "masraf_merkezi_kodu") or gorev_yeri
            ad = al(satir, "masraf_merkezi_adi") or gorev_yeri
            sirket = al(satir, "sirket") or None
            kayitlar.append(
                MasrafMerkezi(
                    gorev_yeri=gorev_yeri,
                    kod=kod,
                    ad=ad,
                    sirket=sirket,
                    aktif=_dogru_mu(al(satir, "aktif")),
                )
            )
        return cls(kayitlar, kaynak=str(hedef), kaynak_var=True)

    # ------------------------------------------------------------------
    # Sorgular
    # ------------------------------------------------------------------

    def coz(self, gorev_yeri: str) -> dict | None:
        """Gorev yerini masraf merkezine cevirir; bulunamazsa None.

        Donen sozluk: gorev_yeri, masraf_merkezi_kodu, masraf_merkezi_adi,
        sirket, aktif.
        """
        metin = _metin(gorev_yeri)
        if not metin:
            return None
        kayit = self._index.get(_anahtar(metin))
        if kayit is None:
            self._bilinmeyenler.add(metin)
            return None
        return kayit.sozluk()

    def tuzel_kisi_mi(self, etiket: str) -> bool:
        """Etiket bir proje degil, tuzel kisi adi mi?

        Kaynak dosyalardaki santiye kolonu ikisini karistirir; karsilastirma
        yaparken ayirt etmek gerekir.
        """
        anahtar = _anahtar(etiket)
        if not anahtar:
            return False
        if anahtar in TUZEL_KISI_ETIKETLERI:
            return True
        # 'RHI 1/3 - RENSTROYDETAL 2/3' gibi paylasimli etiketler.
        parcalar = [p for p in anahtar.replace("+", " ").split() if p.isalpha()]
        return bool(parcalar) and all(
            any(t.startswith(p) or p in t.split() for t in TUZEL_KISI_ETIKETLERI)
            for p in parcalar
        )

    def eksikleri_bildir(self, gorev_yerleri: set) -> list[str]:
        """Haritada karsiligi olmayan gorev yerlerini sirali dondurur."""
        eksik: set[str] = set()
        for deger in gorev_yerleri:
            metin = _metin(deger)
            if not metin:
                continue
            if self._index.get(_anahtar(metin)) is None:
                eksik.add(metin)
        return sorted(eksik)

    @property
    def bilinmeyenler(self) -> list[str]:
        """Calisma sirasinda cozulemeyen gorev yerleri (biriktirilir)."""
        return sorted(self._bilinmeyenler)

    @property
    def kayitlar(self) -> list[MasrafMerkezi]:
        """Haritadaki tum kayitlar."""
        return list(self._kayitlar)

    def kod_adlari(self) -> dict[str, str]:
        """Masraf merkezi kodu -> adi sozlugu (ozet ve cikti icin)."""
        return {k.kod: k.ad for k in self._kayitlar}

    def istatistik(self) -> dict:
        """Haritanin ozeti."""
        return {
            "kayit_sayisi": len(self._kayitlar),
            "aktif_sayisi": sum(1 for k in self._kayitlar if k.aktif),
            "kaynak": self.kaynak,
            "kaynak_var": self.kaynak_var,
            "bilinmeyen_sayisi": len(self._bilinmeyenler),
        }

    def __len__(self) -> int:
        return len(self._kayitlar)


# ----------------------------------------------------------------------
# Cozumleme
# ----------------------------------------------------------------------


def _durum_belirle(guven: float, uyarilar: list[str], masraf_merkezi: str | None,
                   guven_esigi: float, alt_esik: float) -> str:
    """Guven skoru ve uyarilara bakarak cikti sayfasini belirler."""
    if guven < alt_esik or not masraf_merkezi:
        return DURUM_ESLESMEDI
    if guven >= guven_esigi and not uyarilar:
        return DURUM_OTOMATIK
    return DURUM_INCELE


def _sicilsiz_sonuc(
    satir: GiderSatiri,
    eslesme: Eslesme,
    harita: MasrafMerkeziHaritasi,
    ek_masraf_merkezi: str | None,
    guven_esigi: float,
    alt_esik: float,
) -> Sonuc:
    """Personel sicili bulunamayan satirlar icin sonuc uretir.

    Sicil yoksa masraf merkezi personel kaydindan turetilemez. Yine de iki
    yedek kaynak vardir ve ikisi de ONERI olarak yazilir:
        * harici kisiler / ek kisi defterindeki masraf merkezi (kullanicinin
          daha once ogretttigi bilgi)
        * kaynak dosyanin kendi santiye kolonu
    """
    uyarilar: list[str] = []
    masraf_merkezi: str | None = None
    masraf_merkezi_adi: str | None = None
    gorev_yeri: str | None = None
    sirket: str | None = None

    oneri = _metin(ek_masraf_merkezi) or _metin(satir.masraf_merkezi_kaynak)
    kaynagi = "defter" if _metin(ek_masraf_merkezi) else "kaynak dosya"

    if oneri:
        cozum = harita.coz(oneri)
        if cozum:
            gorev_yeri = cozum["gorev_yeri"]
            masraf_merkezi = cozum["masraf_merkezi_kodu"]
            masraf_merkezi_adi = cozum["masraf_merkezi_adi"]
            sirket = cozum["sirket"]
            if not cozum["aktif"]:
                uyarilar.append(
                    f"Masraf merkezi '{masraf_merkezi}' haritada pasif isaretli."
                )
        elif harita.tuzel_kisi_mi(oneri):
            uyarilar.append(
                f"Kaynak dosyada '{oneri}' yaziyor; bu bir tuzel kisi adi, proje degil. "
                "Masraf merkezi belirlenemedi."
            )
        else:
            masraf_merkezi = oneri
            masraf_merkezi_adi = oneri
            uyarilar.append(
                f"'{oneri}' masraf merkezi haritasinda tanimli degil; "
                f"{kaynagi} degeri oldugu gibi kullanildi."
            )

    if masraf_merkezi is None:
        if oneri:
            uyarilar.append(
                "Personel sicili bulunamadi ve eldeki santiye bilgisinden masraf "
                "merkezi turetilemedi; elle atanmali."
            )
        else:
            uyarilar.append(
                "Personel sicili bulunamadi ve kaynak dosyada da santiye bilgisi yok; "
                "masraf merkezi elle atanmali."
            )
    elif eslesme.yontem == "yok":
        uyarilar.append(
            f"Kisi eslesmedi; masraf merkezi '{masraf_merkezi}' {kaynagi} bilgisinden "
            "ONERI olarak yazildi, dogrulayin."
        )
    else:
        uyarilar.append(
            "Kisi personel ana verisinde degil; masraf merkezi defterden/kaynak "
            "dosyadan alindi, dogrulayin."
        )

    durum = _durum_belirle(eslesme.guven, uyarilar, masraf_merkezi, guven_esigi, alt_esik)
    # Sicili olmayan ve hicbir oneri bulunamayan satir her zaman ESLESMEDI'dir.
    if masraf_merkezi is None:
        durum = DURUM_ESLESMEDI

    satir.ek["masraf_merkezi_adi"] = masraf_merkezi_adi or ""
    satir.ek["cozum_kaynagi"] = kaynagi if masraf_merkezi else "yok"

    return Sonuc(
        satir=satir,
        eslesme=eslesme,
        donem=None,
        gorev_yeri=gorev_yeri,
        masraf_merkezi=masraf_merkezi,
        sirket=sirket,
        sirket2=None,
        statu=None,
        kategori=None,
        cikis_tarihi=None,
        durum=durum,
        uyarilar=uyarilar,
    )


def masraf_merkezi_coz(
    satir: GiderSatiri,
    eslesme: Eslesme,
    defter: Any,
    harita: MasrafMerkeziHaritasi,
    *,
    guven_esigi: float = GUVEN_ESIGI,
    alt_esik: float = ALT_ESIK,
    ek_masraf_merkezi: str | None = None,
    tuzel_kisi_uyar: bool = False,
    son_donem: date | None = None,
) -> Sonuc:
    """Bir gider satirini nihai masraf merkezine baglar.

    Args:
        satir: Kaynak dosyadan cikarilmis gider satiri.
        eslesme: Eslestiricinin urettigi kisi eslesmesi.
        defter: ``kayit.PersonelDefteri`` ornegi (tip bagimliligi olmasin diye
            ``Any``; ``donem_kaydi`` ve ``sicil_ile`` metotlari kullanilir).
        harita: Gorev yeri -> masraf merkezi haritasi.
        guven_esigi: Bu ve ustu guven + uyarisiz satir OTOMATIK sayilir.
        alt_esik: Bu esigin altindaki guven ESLESMEDI sayilir.
        ek_masraf_merkezi: Sicili olmayan kisiler icin defterden gelen oneri.
        tuzel_kisi_uyar: Kaynak dosyadaki TUZEL KISI etiketi personel
            kaydindaki 'Sirket 2' ile celisirse uyari uretilsin mi?
        son_donem: Personel dosyasindaki en son snapshot donemi. Verilirse
            belge tarihi bu donemden sonraki bir aya dusuyorsa uyari eklenir.

    Returns:
        ``Sonuc``. ``satir.ek`` sozlugune 'masraf_merkezi_adi' ve
        'cozum_kaynagi' anahtarlari yazilir (cikti modulu bunlari okur).
    """
    if not eslesme.sicil:
        return _sicilsiz_sonuc(
            satir, eslesme, harita, ek_masraf_merkezi, guven_esigi, alt_esik
        )

    uyarilar: list[str] = []
    belge_tarihi = _gecerli_tarih(satir.belge_tarihi)
    kayit = defter.donem_kaydi(eslesme.sicil, belge_tarihi)
    if kayit is None:
        kayit = defter.sicil_ile(eslesme.sicil)
        if kayit is not None:
            uyarilar.append(
                f"Sicil {eslesme.sicil} icin donem kaydi bulunamadi; en guncel kayit kullanildi."
            )
    if kayit is None:
        uyarilar.append(
            f"Sicil {eslesme.sicil} personel ana verisinde bulunamadi. "
            "Alias defterindeki sicil eski veya hatali olabilir."
        )
        satir.ek["masraf_merkezi_adi"] = ""
        satir.ek["cozum_kaynagi"] = "yok"
        return Sonuc(
            satir=satir,
            eslesme=eslesme,
            donem=None,
            gorev_yeri=None,
            masraf_merkezi=None,
            sirket=None,
            sirket2=None,
            statu=None,
            kategori=None,
            cikis_tarihi=None,
            durum=DURUM_ESLESMEDI,
            uyarilar=uyarilar,
        )

    donem: date | None = _gecerli_tarih(kayit.get("donem"))
    gorev_yeri: str | None = kayit.get("gorev_yeri")
    sirket2: str | None = kayit.get("sirket2")
    statu: str | None = kayit.get("statu")
    kategori: str | None = kayit.get("kategori")
    cikis_tarihi: date | None = _gecerli_tarih(kayit.get("cikis_tarihi"))

    # 1) Donem secimi guvenilir mi?
    if kayit.get("_donem_tahmini"):
        if belge_tarihi is None:
            uyarilar.append(
                "Belge tarihi okunamadi; personel kaydinin en guncel donemi kullanildi."
            )
        else:
            uyarilar.append(
                f"Belge tarihi ({_tarih_metni(belge_tarihi)}) bu kisinin personel "
                f"kayitlarindaki donemlerin disinda; en yakin donem "
                f"({_tarih_metni(donem)}) kullanildi (yeni giren olabilir)."
            )
    elif (
        son_donem is not None
        and belge_tarihi is not None
        and _ay_basi_sonrasi(belge_tarihi, son_donem)
    ):
        uyarilar.append(
            f"Belge tarihi ({_tarih_metni(belge_tarihi)}) personel dosyasindaki "
            f"son donemden ({_tarih_metni(son_donem)}) sonra; personel dosyasini guncelleyin."
        )

    # 2) Kisi belge tarihinden once isten ayrilmis mi?
    if (
        kategori == "Cikis"
        and cikis_tarihi is not None
        and belge_tarihi is not None
        and cikis_tarihi < belge_tarihi
    ):
        uyarilar.append(
            f"Kisi belge tarihinden once isten ayrilmis (cikis: {_tarih_metni(cikis_tarihi)}). "
            "Son calistigi proje kullanildi."
        )

    # 3) Gorev yeri -> masraf merkezi
    masraf_merkezi: str | None = None
    masraf_merkezi_adi: str | None = None
    harita_sirketi: str | None = None
    if gorev_yeri:
        cozum = harita.coz(gorev_yeri)
        if cozum:
            masraf_merkezi = cozum["masraf_merkezi_kodu"]
            masraf_merkezi_adi = cozum["masraf_merkezi_adi"]
            harita_sirketi = cozum["sirket"]
            if not cozum["aktif"]:
                uyarilar.append(
                    f"Masraf merkezi '{masraf_merkezi}' haritada pasif isaretli; "
                    "gecerli kodu kontrol edin."
                )
        else:
            masraf_merkezi = gorev_yeri
            masraf_merkezi_adi = gorev_yeri
            uyarilar.append(
                f"'{gorev_yeri}' masraf merkezi haritasinda tanimli degil; gorev yeri "
                f"oldugu gibi kullanildi. {DOSYA_ADI} dosyasina ekleyin."
            )
    else:
        uyarilar.append(
            "Personel kaydinda gorev yeri bos; masraf merkezi belirlenemedi."
        )

    # 4) Kaynak dosyadaki santiye etiketi ile karsilastirma
    kaynak_etiket = _metin(satir.masraf_merkezi_kaynak)
    if kaynak_etiket and masraf_merkezi:
        if harita.tuzel_kisi_mi(kaynak_etiket):
            if tuzel_kisi_uyar and sirket2 and _anahtar(kaynak_etiket) != _anahtar(sirket2):
                uyarilar.append(
                    f"Kaynak dosyada tuzel kisi '{kaynak_etiket}' yaziyor, personel "
                    f"kaydina gore '{sirket2}'. Kontrol edin."
                )
        else:
            kaynak_cozum = harita.coz(kaynak_etiket)
            kaynak_kodu = kaynak_cozum["masraf_merkezi_kodu"] if kaynak_cozum else None
            if kaynak_kodu is None:
                if _anahtar(kaynak_etiket) != _anahtar(gorev_yeri):
                    uyarilar.append(
                        f"Kaynak dosyada '{kaynak_etiket}' yaziyor; bu deger masraf merkezi "
                        f"haritasinda tanimli degil. Personel kaydina gore '{masraf_merkezi}'. "
                        "Kontrol edin."
                    )
            elif kaynak_kodu != masraf_merkezi:
                uyarilar.append(
                    f"Kaynak dosyada '{kaynak_etiket}' ({kaynak_kodu}) yaziyor, personel "
                    f"kaydina gore '{masraf_merkezi}'. Kontrol edin."
                )

    durum = _durum_belirle(eslesme.guven, uyarilar, masraf_merkezi, guven_esigi, alt_esik)

    satir.ek["masraf_merkezi_adi"] = masraf_merkezi_adi or ""
    satir.ek["cozum_kaynagi"] = "personel"

    return Sonuc(
        satir=satir,
        eslesme=eslesme,
        donem=donem,
        gorev_yeri=gorev_yeri,
        masraf_merkezi=masraf_merkezi,
        sirket=harita_sirketi or sirket2,
        sirket2=sirket2,
        statu=statu,
        kategori=kategori,
        cikis_tarihi=cikis_tarihi,
        durum=durum,
        uyarilar=uyarilar,
    )
