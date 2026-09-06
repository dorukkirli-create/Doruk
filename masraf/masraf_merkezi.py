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
Bazi kaynak dosyalarda insan tarafindan doldurulmus bir santiye kolonu bulunur
(``GiderSatiri.masraf_merkezi_kaynak``). ISLETME KARARI: bu kolon PROJE olarak
yorumlanir ve personel kaydindan cozulen proje ile karsilastirilir.

Pratikte kolon iki tur deger tasiyor:

* PROJE adi ('GPP Proje', 'Udokan GMK', 'AMURSKI GAZ ISLETME FABRIKASI').
  Bunlar ``EK_ESANLAMLILAR`` uzerinden personel verisindeki 'Gorev Yeri'
  degerlerine cevrilir ve cozulen proje ile karsilastirilir. Farklilarsa
  satir ``PROJE UYUSMAZLIGI`` uyarisi alir ve incelemeye duser. Elle yapilan
  atama hatalarini yakalayan grup budur.

* TUZEL KISI adi ('RHI', 'RENSERVIS', 'ONE TOWER', 'RC PETER'). Bunlar proje
  degildir; o satirda karsilastirilacak proje bilgisi YOKTUR. Celiski de
  yoktur, bu yuzden satir incelemeye DUSURULMEZ. Satir
  ``ek['kaynak_proje_yerine_sirket']`` ile isaretlenir ve ozetteki
  ``kaynak_proje_karsilastirmasi['proje_yok']`` sayacinda toplu raporlanir.
  Boylece finans ekibi kaynak dosyalarda kac satirda proje yerine sirket
  yazildigini gorur, ama operator ayni bilgiyi yuzlerce kez elle onaylamaz.

Olculen sonuc (Temmuz 2026, tek Outlook mesaji, 405 satir):
uyusan 109, uyusmayan 4, proje yerine sirket yazilmis 98, etiketsiz 194.

Tuzel kisi etiketinin personel kaydindaki 'Sirket 2' ile karsilastirilmasi
ayri bir konudur ve ``tuzel_kisi_uyar=True`` ile acilabilir; varsayilan
kapalidir cunku elle etiketleme sirket duzeyinde kendi icinde tutarsizdir.
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
    # Haritanin kendi sirket kodlari da tuzel kisidir.
    "RSS",
    "BSK",
    "BSA",
    "ULTK",
    "RC",
})

#: Tuzel kisi etiketlerini haritanin sirket koduna ceviren sozluk. Kaynak
#: dosyada 'RENSTROYDETAL 2/3' yazan pay, Sirket Kirilimi'nde 'RSS' altinda
#: toplanmali; aksi halde ayni tuzel kisi iki satir olur (olculdu: 134,28 USD).
SIRKET_KANONIK: dict[str, str] = {
    "RENSTROYDETAL": "RSS", "RENSERVIS": "RSS", "RSD": "RSS", "RS": "RSS", "RSS": "RSS",
    "RC PETER": "RC", "RC PETERSBURG": "RC", "RC MOSKOVA": "RC", "RC MOSCOW": "RC",
    "ONE TOWER": "RC", "TOP TOWER": "RC", "RC": "RC",
    "UST LUGA": "UST LUGA", "USTLUGA": "UST LUGA", "ULTK": "UST LUGA",
    "RHI": "RHI", "RHI RUSSIA": "RHI",
    "BSK": "BSK", "BSK MANAGEMENT": "BSK", "BSK MANAGEMENT GROUP": "BSK",
    "BSA": "BSA", "SAREN": "SAREN", "KRONDEX": "KRONDEX", "YAKA LLC": "YAKA LLC", "YAKA": "YAKA LLC",
}


def sirket_kanonik(etiket: Any) -> str | None:
    """Bir tuzel kisi etiketini haritanin sirket koduna cevirir; bilinmiyorsa None."""
    anahtar = _anahtar(etiket)
    if not anahtar:
        return None
    return SIRKET_KANONIK.get(anahtar)

#: Kaynak dosyalarda gecen proje yazimlarinin personel verisindeki 'Gorev Yeri'
#: karsiliklari. Kullanici CSV'sini kirletmemek icin kod tarafinda tutulur;
#: CSV'ye ayni gorev yeri icin ikinci bir satir eklenmesi de calisir.
EK_ESANLAMLILAR: dict[str, str] = {
    "GPP PROJE": "GPP Project",
    "GPP PROJESI": "GPP Project",
    "UST LUGA GPP": "GPP Project",
    "UST LUGA GPP PROJESI": "GPP Project",
    "UST LUGA GAS PROCESSING COMPLEX GPP": "GPP Project",
    # Olculdu: saglik kontrol listesinde bu etiketi tasiyan 7 kisiden 6'si
    # personel verisinde "GPP Project" gorunuyor. Etiket fabrikayi degil
    # kompleksin tamamini kastediyor.
    "UST LUGA GAS PROCESSING COMPLEX GPC": "GPP Project",
    "UST LUGA GAS PROCESSING COMPLEX": "GPP Project",
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
    metin = str(deger)
    # Tireler ascii_katla'dan ONCE bosluga cevrilir: en/em-dash ASCII'de
    # yoktur ve katlama onlari siler ('A–B' -> 'AB' olur, olculdu).
    for karakter in _TIRE_KARAKTERLERI:
        metin = metin.replace(karakter, " ")
    metin = ascii_katla(metin).upper()
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


_AYLAR = ("Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
          "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik")


def _ay_metni(deger) -> str:
    """Tarihi 'Temmuz 2026' bicimine cevirir; gider ayini okunakli yazmak icin."""
    t = _gecerli_tarih(deger)
    if t is None:
        return "?"
    return f"{_AYLAR[t.month - 1]} {t.year}"


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
                 kaynak: str | None = None, kaynak_var: bool = True,
                 yukleme_uyarisi: str | None = None) -> None:
        self.kaynak = kaynak
        self.kaynak_var = kaynak_var
        #: Dosya var ama kolonlari eksik/bozuk ise burada aciklanir. Bos harita
        #: 'gecerli' sayilirsa kullanici 'harita bozuk' yerine 'N gorev yeri
        #: tanimsiz' uyarisi okur ve yanlis yere ugrasir (olculdu).
        self.yukleme_uyarisi = yukleme_uyarisi
        self._kayitlar: list[MasrafMerkezi] = list(kayitlar or ())
        self._index: dict[str, MasrafMerkezi] = {}
        self._bilinmeyenler: set[str] = set()
        self._indeksle()

    # ------------------------------------------------------------------
    # Kurulum
    # ------------------------------------------------------------------

    def _indeksle(self) -> None:
        """Arama indeksini kurar. Once gorev yerleri, sonra kod ve adlar.

        Gorev yeri anahtari her zaman kazanir: bir satirin kodu baska satirin
        gorev yeriyle cakisirsa gorev yeri sahibi bulunmali (golgeleme yok).
        """
        self._index.clear()
        for kayit in self._kayitlar:
            anahtar = _anahtar(kayit.gorev_yeri)
            if anahtar:
                self._index.setdefault(anahtar, kayit)
        for kayit in self._kayitlar:
            for aday in (kayit.kod, kayit.ad):
                anahtar = _anahtar(aday)
                if anahtar:
                    self._index.setdefault(anahtar, kayit)
        # Esanlamlilar: hedef gorev yeri haritada varsa onun kaydina baglanir.
        for yazim, hedef in EK_ESANLAMLILAR.items():
            kayit = self._index.get(_anahtar(hedef))
            if kayit is not None:
                self._index.setdefault(_anahtar(yazim), kayit)

    def _golgelenenler(self) -> list[str]:
        """Kodu/adi baska satirin gorev yeriyle cakisan kayitlari listeler."""
        gy = {_anahtar(k.gorev_yeri): k for k in self._kayitlar if _anahtar(k.gorev_yeri)}
        cikti: list[str] = []
        for k in self._kayitlar:
            for aday in (k.kod, k.ad):
                a = _anahtar(aday)
                if a and a in gy and gy[a] is not k:
                    cikti.append(f"'{aday}' ({k.gorev_yeri}) ~ gorev yeri '{gy[a].gorev_yeri}'")
        return cikti

    @classmethod
    def yukle(cls, yol: str | Path) -> "MasrafMerkeziHaritasi":
        """CSV harita dosyasini okur.

        Ayirici otomatik bulunur (virgul veya noktali virgul); Excel'in
        Turkce/Rusca yerel ayarla kaydettigi dosyalar da okunur.
        """
        hedef = Path(yol)
        if not hedef.exists():
            return cls([], kaynak=str(hedef), kaynak_var=False)

        kodlama_uyarisi: str | None = None
        try:
            ham = hedef.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            # Turkce Windows Excel'in 'CSV (virgulle ayrilmis)' kaydi cp1254'tur.
            try:
                ham = hedef.read_text(encoding="cp1254")
                kodlama_uyarisi = (
                    f"Masraf merkezi haritasi ({hedef.name}) UTF-8 degil, Windows-1254 "
                    "olarak okundu. Excel'de 'CSV UTF-8' olarak kaydetmeniz onerilir."
                )
            except (OSError, UnicodeDecodeError):
                try:
                    ham = hedef.read_text(encoding="utf-8", errors="replace")
                    kodlama_uyarisi = (
                        f"Masraf merkezi haritasi ({hedef.name}) okunurken bozuk karakterler "
                        "bulundu; Turkce harfli gorev yerleri cozulemeyebilir. Dosyayi UTF-8 kaydedin."
                    )
                except OSError:
                    return cls([], kaynak=str(hedef), kaynak_var=False)
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

        eksik = [k for k in ("gorev_yeri", "masraf_merkezi_kodu") if k not in alan_haritasi]
        if eksik:
            return cls([], kaynak=str(hedef), kaynak_var=True, yukleme_uyarisi=(
                f"Masraf merkezi haritasi ({hedef.name}) okunamadi: baslikta "
                f"{', '.join(eksik)} kolonu yok (bulunan: "
                f"{', '.join(okuyucu.fieldnames or []) or 'hic'}). Dosya bozuk ya da "
                "yanlis kaydedilmis; hicbir gorev yeri cozulemeyecek."))

        kayitlar: list[MasrafMerkezi] = []
        gorulen_gy: dict[str, str] = {}
        tekrarlar: list[str] = []
        for satir in okuyucu:
            gorev_yeri = al(satir, "gorev_yeri")
            if not gorev_yeri:
                continue
            kod = al(satir, "masraf_merkezi_kodu") or gorev_yeri
            ad = al(satir, "masraf_merkezi_adi") or gorev_yeri
            # Sirket kodu normalize: 'ust luga' ile 'UST LUGA' ayni sirkettir.
            sirket = (al(satir, "sirket") or "").strip().upper() or None
            gy_anahtar = _anahtar(gorev_yeri)
            if gy_anahtar in gorulen_gy:
                tekrarlar.append(f"'{gorev_yeri}' (ilk satir: '{gorulen_gy[gy_anahtar]}')")
            else:
                gorulen_gy[gy_anahtar] = gorev_yeri
            kayitlar.append(
                MasrafMerkezi(
                    gorev_yeri=gorev_yeri,
                    kod=kod,
                    ad=ad,
                    sirket=sirket,
                    aktif=_dogru_mu(al(satir, "aktif")),
                )
            )
        uyarilar: list[str] = []
        if kodlama_uyarisi:
            uyarilar.append(kodlama_uyarisi)
        if not kayitlar:
            uyarilar.append(f"Masraf merkezi haritasi ({hedef.name}) bos: basliklar dogru ama "
                            "hicbir satir okunamadi.")
        if tekrarlar:
            # Ayni gorev yeri iki kez yazildiysa ILK satir kazanir; kullanici
            # duzeltmeyi sona eklerse etkisiz kalir. Bunu soylemek gerekir.
            uyarilar.append(
                f"Masraf merkezi haritasinda ayni gorev yeri birden fazla satirda: "
                + "; ".join(tekrarlar[:5]) + ". Ilk satir gecerli sayildi; digerini silin."
            )
        harita = cls(kayitlar, kaynak=str(hedef), kaynak_var=True,
                     yukleme_uyarisi=(" | ".join(uyarilar) if uyarilar else None))
        golgeler = harita._golgelenenler()
        if golgeler:
            harita.yukleme_uyarisi = ((harita.yukleme_uyarisi + " | ") if harita.yukleme_uyarisi else "") + (
                "Masraf merkezi haritasinda bir satirin kodu/adi baska bir satirin gorev yeriyle "
                "cakisiyor, ikinci satir kendi gorev yeriyle bulunamaz: " + "; ".join(golgeler[:5]))
        return harita

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
        # 'RHI 1/3 - RENSTROYDETAL 2/3' gibi paylasimli etiketler: kelimeler
        # bastan sona TAM tuzel kisi etiketleriyle (1-3 kelimelik) ortulmeli.
        # Onek eslemesi ('R', 'U', 'SA') tuzel sayilirdi; kaldirildi.
        parcalar = [p for p in anahtar.replace("+", " ").split() if p.isalpha()]
        if not parcalar:
            return False
        i = 0
        while i < len(parcalar):
            uzunluk = 0
            for n in (3, 2, 1):
                if " ".join(parcalar[i:i + n]) in TUZEL_KISI_ETIKETLERI:
                    uzunluk = n
                    break
            if not uzunluk:
                return False
            i += uzunluk
        return True

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


#: Kimligi ISIMDEN turetilen yontemler; es isimli kontrolu bunlara uygulanir.
_ISIM_TABANLI_YONTEMLER = frozenset({
    "tam_isim", "alt_kume", "prefix", "transliterasyon", "bulanik", "alias",
})


def _yardimci_kaydi(yardimci: Any, sicil: str | None) -> dict | None:
    """1C listesinde sicile karsilik gelen kayit (yoksa None)."""
    if yardimci is None or not sicil:
        return None
    try:
        return yardimci.sicil_ile(str(sicil))
    except Exception:  # noqa: BLE001
        return None


def _ayni_proje(harita: Any, a: Any, b: Any) -> bool:
    """Iki gorev yeri metni ayni masraf merkezine mi cozuluyor?

    Once harita kodu, harita cozemezse ham anahtar karsilastirilir.
    """
    if harita is not None:
        try:
            ca = harita.coz(a) if a else None
            cb = harita.coz(b) if b else None
        except Exception:  # noqa: BLE001
            ca = cb = None
        if ca and cb:
            return ca["masraf_merkezi_kodu"] == cb["masraf_merkezi_kodu"]
    return _anahtar(a or "") == _anahtar(b or "")


def _yardimci_es_isimliler(yardimci: Any, eslesme: Any, gorev_yeri: str | None,
                           harita: Any = None) -> list[tuple[str, str]]:
    """1C listesinde ayni isimli, FARKLI sicilli ve FARKLI projedeki kisiler.

    Returns:
        [(gorev_yeri, sirket)] listesi; bos ise es isimli yok.
    """
    if yardimci is None:
        return []
    try:
        from masraf.metin import isim_normalize
        ad = eslesme.ad_soyad or ""
        adaylar = yardimci.isimle_adaylar(isim_normalize(ad)) if ad else []
    except Exception:  # noqa: BLE001
        return []
    sonuc: list[tuple[str, str]] = []
    for sicil in adaylar:
        if str(sicil) == str(eslesme.sicil):
            continue
        kayit = _yardimci_kaydi(yardimci, sicil)
        if not kayit or not kayit.get("gorev_yeri"):
            continue
        if _ayni_proje(harita, kayit["gorev_yeri"], gorev_yeri):
            continue
        sonuc.append((str(kayit["gorev_yeri"]), str(kayit.get("sirket2") or kayit.get("sirket") or "?")))
    return sonuc


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
        donem_eslesme="yok",
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
    yardimci: Any = None,
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
    if kayit is None and yardimci is not None:
        # Ana veride yok: 1C personel listesine bak. Grup sirketleri
        # (Renservis, Renstroydetal, RC, One Tower, Top Tower) ancak orada.
        kayit = yardimci.donem_kaydi(eslesme.sicil, belge_tarihi)
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
            donem_eslesme="yok",
        )

    donem: date | None = _gecerli_tarih(kayit.get("donem"))
    gorev_yeri: str | None = kayit.get("gorev_yeri")
    sirket2: str | None = kayit.get("sirket2")
    statu: str | None = kayit.get("statu")
    kategori: str | None = kayit.get("kategori")
    cikis_tarihi: date | None = _gecerli_tarih(kayit.get("cikis_tarihi"))

    # 1) Gider ayi ile donem ayi ortusuyor mu?
    #    Kural: masraf, giderin YAPILDIGI AYDAKI kayda gore mahsuplasir.
    #    Kisinin projesi ve masraf merkezi aydan aya degisebilir.
    donem_eslesme = kayit.get("_donem_eslesme")
    if donem_eslesme == "yardimci_defter":
        firma = kayit.get("sirket")
        grup = kayit.get("sirket2")
        sirket_metni = (f"sirket {grup}" + (f", 1C firmasi {firma}" if firma and firma != grup else "")
                        if grup else f"sirket {firma}")
        uyarilar.append(
            f"Kisi ana personel verisinde yok, 1C personel listesinden alindi "
            f"({sirket_metni}, proje {kayit.get('gorev_yeri')}). "
            "1C listesi tek tarihli oldugu icin gider ayindaki durum dogrulanamadi."
        )
    elif donem_eslesme == "tarihsiz":
        uyarilar.append(
            "Belge tarihi okunamadi; personel kaydinin en guncel donemi kullanildi. "
            "Gider ayi dogrulanamadi."
        )
    elif donem_eslesme == "ilk_donem_oncesi":
        ise_giris = _gecerli_tarih(kayit.get("ise_giris_tarihi"))
        if ise_giris is not None and belge_tarihi is not None and ise_giris <= belge_tarihi:
            # Kisi zaten calisiyordu; personel dosyasi o ayi kapsamiyor.
            uyarilar.append(
                f"Gider tarihi ({_tarih_metni(belge_tarihi)}) personel dosyasinin bu kisi icin "
                f"kapsadigi ilk donemden ({_tarih_metni(donem)}) once; kisi {_tarih_metni(ise_giris)} "
                "tarihinden beri calisiyor, dosya o ayi icermiyor. Ilk donemin masraf merkezi "
                "kullanildi; o aydaki projesi farkli olabilir."
            )
        else:
            uyarilar.append(
                f"Gider tarihi ({_tarih_metni(belge_tarihi)}) kisinin ilk personel "
                f"kaydindan ({_tarih_metni(donem)}) ONCE. Kisi o tarihte henuz ise "
                "baslamamis; mobilizasyon veya aday seyahati olabilir. Ilk donemin "
                "masraf merkezi kullanildi."
            )
    elif donem_eslesme == "onceki_donem":
        dosya_eski = (
            son_donem is not None and belge_tarihi is not None
            and _ay_basi_sonrasi(belge_tarihi, son_donem)
        )
        if dosya_eski:
            # Gider ayi personel dosyasinin son doneminden SONRA: kimsenin o
            # ayda kaydi olamaz. Sorun kisi degil, dosyanin guncel olmamasi.
            donem_eslesme = "personel_dosyasi_eski"
            mesaj = (
                f"Personel dosyasi {_ay_metni(son_donem)} ile bitiyor; gider ayinin "
                f"({_ay_metni(belge_tarihi)}) kaydi henuz yok. Kisinin son donemi "
                f"({_ay_metni(donem)}) kullanildi. Personel dosyasini guncelleyip tekrar calistirin."
            )
        elif kategori == "Cikis":
            mesaj = (
                f"Kisinin gider ayinda ({_ay_metni(belge_tarihi)}) personel kaydi yok; "
                f"isten ayrilmis. Kayitli son donemi {_ay_metni(donem)}; o donemin masraf "
                "merkezi kullanildi. Cikis masrafi ise dogru santiyedir, kontrol edin."
            )
        else:
            mesaj = (
                f"Kisinin gider ayinda ({_ay_metni(belge_tarihi)}) personel kaydi yok ve "
                f"cikis kaydi da yok; kayit {_ay_metni(donem)} donemiyle kesilmis. O donemin "
                "masraf merkezi kullanildi; kisi baska sirkete/projeye gecmis olabilir, kontrol edin."
            )
        # Ana veriden cikmis ama 1C listesinde BASKA bir projede aktifse,
        # muhtemelen grup sirketine gecmistir; masraf oraya ait olabilir.
        # Karsilastirma harita KODU uzerinden yapilir: 'Reshetnikova Office'
        # ile 'St. Petersburg Office' ayni projedir (olculdu: 2 yanlis alarm).
        baska = _yardimci_kaydi(yardimci, eslesme.sicil)
        if baska and baska.get("gorev_yeri") and not _ayni_proje(harita, baska["gorev_yeri"], gorev_yeri):
            mesaj += (
                f" DIKKAT: 1C listesinde kisi '{baska['gorev_yeri']}' "
                f"({baska.get('sirket2') or baska.get('sirket') or '?'}) projesinde gorunuyor; "
                "grup sirketine gecmis olabilir, masraf oraya ait olabilir."
            )
        uyarilar.append(mesaj)

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
    # Masraf merkezi haritada tanimli bir finans kodu mu, yoksa cozulemedigi
    # icin oldugu gibi tasinan ham gorev yeri metni mi? Mahsuplasma tablosu
    # ikisini ayirmak zorundadir: ilki muhasebeye gidebilir, ikincisi once
    # haritaya eklenmelidir.
    satir.ek["masraf_merkezi_haritada"] = False
    if gorev_yeri:
        cozum = harita.coz(gorev_yeri)
        if cozum:
            masraf_merkezi = cozum["masraf_merkezi_kodu"]
            masraf_merkezi_adi = cozum["masraf_merkezi_adi"]
            harita_sirketi = cozum["sirket"]
            satir.ek["masraf_merkezi_haritada"] = True
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
            # Kaynak dosyada proje yerine SIRKET adi yazilmis. Bu satirda proje
            # bilgisi yoktur; celiski de yoktur. Satiri incelemeye dusurmez,
            # sadece isaretlenir ve ozette toplu olarak raporlanir.
            satir.ek["kaynak_proje_yerine_sirket"] = kaynak_etiket
            if tuzel_kisi_uyar and sirket2 and _anahtar(kaynak_etiket) != _anahtar(sirket2):
                uyarilar.append(
                    f"Kaynak dosyada tuzel kisi '{kaynak_etiket}' yaziyor, personel "
                    f"kaydina gore '{sirket2}'. Kontrol edin."
                )
        else:
            kaynak_cozum = harita.coz(kaynak_etiket)
            kaynak_kodu = kaynak_cozum["masraf_merkezi_kodu"] if kaynak_cozum else None
            if kaynak_kodu is None:
                # Cozulen deger kaynak etiketin kendisiyse (haritada karsiligi
                # olmadigi icin oldugu gibi kullanilmis) uyusmazlik yoktur.
                ayni = (_anahtar(kaynak_etiket) == _anahtar(gorev_yeri)
                        or _anahtar(kaynak_etiket) == _anahtar(masraf_merkezi))
                if not ayni:
                    uyarilar.append(
                        f"Kaynak dosyada proje '{kaynak_etiket}' yaziyor; bu deger masraf "
                        f"merkezi haritasinda tanimli degil. Personel kaydina gore "
                        f"'{masraf_merkezi}'. Kontrol edin."
                    )
            elif kaynak_kodu != masraf_merkezi:
                uyarilar.append(
                    f"PROJE UYUSMAZLIGI: kaynak dosyada '{kaynak_etiket}' ({kaynak_kodu}) "
                    f"yaziyor, personel kaydina gore '{masraf_merkezi}'. Kontrol edin."
                )

    # 5) Isimle bulunan kisinin 1C listesinde AYNI ISIMLI ama farkli projede
    #    baska bir kaydi varsa otomatik kabul edilmez. Olculdu: Temmuz 2026'da
    #    iki satir boyle; ikisi de ULF-GPC-RHI yerine RSS Lytkarino olabilir.
    if eslesme.yontem in _ISIM_TABANLI_YONTEMLER and eslesme.sicil:
        esler = _yardimci_es_isimliler(yardimci, eslesme, gorev_yeri, harita)
        if esler:
            uyarilar.append(
                "1C listesinde ayni isimli baska kisi var: "
                + "; ".join(f"{y} ({s})" for y, s in esler[:2])
                + ". Dogru kisi olduguna emin olun."
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
        donem_eslesme=donem_eslesme or "yok",
    )
